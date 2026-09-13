/* 巴菲特投资智慧 · Windows 启动器（编译版，替代易被杀软误报的 VBS 脚本）
 * ======================================================================
 * 行为（与 Mac 版 main.swift 对齐）：
 *   1. 扫描 127.0.0.1:8666-8685 是否已有本应用服务
 *      （/api/llm-config 含 buffett-wisdom 标记）；
 *   2. 未运行则拉起内置 Python 本地服务，启动方式三级兜底：
 *        a. 隐藏创建（CREATE_NO_WINDOW）——常规机器无任何窗口；
 *        b. 可见最小化控制台——部分安全软件拦截"隐藏启动"，但放行可见进程；
 *        c. ShellExecute 经由资源管理器路径启动——行为最接近手动双击。
 *      首选端口 8666，被占用自动回退 8667+；
 *   3. 打开 Edge 应用模式窗口承载应用（找不到 Edge 则回退默认浏览器）。
 * 服务 PID 写入 %APPDATA%\巴菲特投资智慧\server.pid（ASCII），供停止/卸载使用。
 * 失败时：错误码 + 安装目录 + python 输出（管道捕获，launch-debug.log）一并提示；
 * 若日志中能解析出服务实际地址（探测异常时），仍会直接打开该地址。
 *
 * 编译（macOS 交叉编译，由 package_windows.py 调用）：
 *   x86_64-w64-mingw32-gcc -O2 -municode -mwindows \
 *       -o 巴菲特投资智慧.exe launcher.c -lws2_32 -lshell32 -luser32 -ladvapi32
 */
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <shellapi.h>
#include <time.h>
#include <stdio.h>
#include <wchar.h>

#define START_PORT 8666
#define MAX_PORT   8685
#define MARKER_ASC "buffett-wisdom"   /* 服务标记（ASCII，字节比较） */

static wchar_t g_appdir[MAX_PATH];     /* 本 exe 所在目录（安装目录） */
static wchar_t g_pidfile[MAX_PATH];    /* %APPDATA%\巴菲特投资智慧\server.pid */
static wchar_t g_debuglog[MAX_PATH];   /* %APPDATA%\巴菲特投资智慧\launch-debug.log */
static wchar_t g_flagfile[MAX_PATH];   /* %APPDATA%\巴菲特投资智慧\server-blocked.flag */
static HANDLE  g_hOutR = NULL;         /* python 输出管道读端（诊断用） */
static HANDLE  g_hProc = NULL;         /* python 进程句柄 */

/* ---------- 通用 ---------- */
static void die_msg(const wchar_t *title, const wchar_t *fmt, ...)
{
    wchar_t buf[2048];
    va_list ap;
    va_start(ap, fmt);
    _vsnwprintf(buf, 2047, fmt, ap);
    va_end(ap);
    MessageBoxW(NULL, buf, title, MB_OK | MB_ICONERROR);
    ExitProcess(1);
}

static void init_paths(void)
{
    wchar_t appdata[MAX_PATH] = L"";
    wchar_t dir[MAX_PATH];
    DWORD n;

    GetModuleFileNameW(NULL, g_appdir, MAX_PATH);
    wchar_t *p = wcsrchr(g_appdir, L'\\');
    if (p) *p = 0;

    n = GetEnvironmentVariableW(L"APPDATA", appdata, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) {
        GetEnvironmentVariableW(L"USERPROFILE", appdata, MAX_PATH);
        wcsncat(appdata, L"\\AppData\\Roaming", MAX_PATH - wcslen(appdata) - 1);
    }
    swprintf(dir, MAX_PATH, L"%ls\\巴菲特投资智慧", appdata);
    CreateDirectoryW(dir, NULL);            /* %APPDATA% 已存在，只建一层 */
    swprintf(g_pidfile, MAX_PATH, L"%ls\\server.pid", dir);
    swprintf(g_debuglog, MAX_PATH, L"%ls\\launch-debug.log", dir);
    swprintf(g_flagfile, MAX_PATH, L"%ls\\server-blocked.flag", dir);
}

/* ---------- 端口探测：GET /api/llm-config 并检查应用标记 ---------- */
static int probe_port(int port)
{
    SOCKET s;
    struct sockaddr_in sa;
    u_long nb = 1;
    fd_set wf;
    struct timeval tv;
    const char *req = "GET /api/llm-config HTTP/1.0\r\nHost: 127.0.0.1\r\n\r\n";
    char buf[4096];
    int got = 0, r;
    DWORD to = 1500;

    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (s == INVALID_SOCKET) return 0;
    memset(&sa, 0, sizeof sa);
    sa.sin_family = AF_INET;
    sa.sin_port = htons((u_short)port);
    sa.sin_addr.s_addr = inet_addr("127.0.0.1");

    ioctlsocket(s, FIONBIO, &nb);
    if (connect(s, (struct sockaddr *)&sa, sizeof sa) != 0) {
        FD_ZERO(&wf); FD_SET(s, &wf);
        tv.tv_sec = 1; tv.tv_usec = 500000;
        if (select(0, NULL, &wf, NULL, &tv) <= 0) { closesocket(s); return 0; }
    }
    /* 关键：连接建立后必须恢复阻塞模式——否则 recv 立即返回 WSAEWOULDBLOCK，
     * 探测读不到响应、还会在服务器写完前关闭连接（WinError 10053） */
    nb = 0;
    ioctlsocket(s, FIONBIO, &nb);
    send(s, req, (int)strlen(req), 0);
    setsockopt(s, SOL_SOCKET, SO_RCVTIMEO, (const char *)&to, sizeof to);
    for (;;) {
        r = recv(s, buf + got, (int)sizeof buf - 1 - got, 0);
        if (r <= 0) break;
        got += r;
        if (got >= (int)sizeof buf - 1) break;
    }
    buf[got] = 0;
    closesocket(s);
    return strstr(buf, MARKER_ASC) != NULL;
}

static int find_existing_server(void)
{
    int p;
    for (p = START_PORT; p <= MAX_PORT; p++)
        if (probe_port(p)) return p;
    return 0;
}

static int wait_server(void)
{
    int t, p;
    for (t = 0; t < 6; t++) {                     /* 首选端口（正常启动路径，3 秒） */
        if (probe_port(START_PORT)) return START_PORT;
        Sleep(500);
    }
    for (t = 0; t < 5; t++) {                     /* 回退端口扫描（4 秒） */
        for (p = START_PORT + 1; p <= MAX_PORT; p++)
            if (probe_port(p)) return p;
        Sleep(800);
    }
    return 0;
}

/* ---------- 启动内置 Python（CreateProcess；visible=1 时最小化可见控制台） ---------- */
static int start_python(int visible)
{
    wchar_t py[MAX_PATH], script[MAX_PATH], cmd[1200];
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    SECURITY_ATTRIBUTES sa;
    HANDLE hOutW = NULL, hNul = NULL;
    wchar_t pidbuf[32];
    char ascii[32];
    DWORD written;
    DWORD flags;

    swprintf(py, MAX_PATH, L"%ls\\python\\python.exe", g_appdir);
    swprintf(script, MAX_PATH, L"%ls\\app\\serve_buffett_app.py", g_appdir);
    if (GetFileAttributesW(py) == INVALID_FILE_ATTRIBUTES)
        die_msg(L"巴菲特投资智慧", L"安装目录缺少文件：\n  %ls\n\n"
                L"可能原因：安全软件拦截了安装解压，或安装包下载不完整。\n"
                L"建议：重新下载安装包后重装（可暂时退出杀毒软件）。", py);
    if (GetFileAttributesW(script) == INVALID_FILE_ATTRIBUTES)
        die_msg(L"巴菲特投资智慧", L"安装目录缺少文件：\n  %ls", script);

    swprintf(cmd, 1200, L"\"%ls\" -u \"%ls\" --no-browser --port %d",
             py, script, START_PORT);

    memset(&si, 0, sizeof si);
    si.cb = sizeof si;
    memset(&sa, 0, sizeof sa);
    sa.nLength = sizeof sa;
    sa.bInheritHandle = TRUE;

    /* stdout/stderr → 管道（诊断日志）；stdin → NUL（避免无效句柄问题） */
    if (CreatePipe(&g_hOutR, &hOutW, &sa, 0))
        SetHandleInformation(g_hOutR, HANDLE_FLAG_INHERIT, 0);
    hNul = CreateFileW(L"NUL", GENERIC_READ,
                       FILE_SHARE_READ | FILE_SHARE_WRITE, &sa,
                       OPEN_EXISTING, 0, NULL);
    si.dwFlags = STARTF_USESTDHANDLES;
    if (visible) {
        si.dwFlags |= STARTF_USESHOWWINDOW;
        si.wShowWindow = SW_SHOWMINIMIZED;
    }
    si.hStdInput = hNul;
    si.hStdOutput = hOutW;
    si.hStdError = hOutW;

    flags = visible ? 0 : CREATE_NO_WINDOW;
    if (!CreateProcessW(NULL, cmd, NULL, NULL, TRUE, flags,
                        NULL, g_appdir, &si, &pi)) {
        if (hOutW) CloseHandle(hOutW);
        if (hNul) CloseHandle(hNul);
        return 0;
    }
    if (hOutW) CloseHandle(hOutW);
    if (hNul) CloseHandle(hNul);
    g_hProc = pi.hProcess;
    CloseHandle(pi.hThread);

    /* 记录 PID（ASCII 文本，供停止服务/卸载器读取） */
    swprintf(pidbuf, 32, L"%lu", pi.dwProcessId);
    int len = WideCharToMultiByte(CP_ACP, 0, pidbuf, -1, ascii, 32, NULL, NULL);
    HANDLE h = CreateFileW(g_pidfile, GENERIC_WRITE, 0, NULL,
                           CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (h != INVALID_HANDLE_VALUE) {
        if (len > 0) WriteFile(h, ascii, (DWORD)(len - 1), &written, NULL);
        CloseHandle(h);
    }
    return 1;
}

/* ---------- 启动兜底：ShellExecute（最接近手动双击，安全软件最不会拦） ---------- */
static int start_python_shell(void)
{
    wchar_t py[MAX_PATH], script[MAX_PATH], args[1200];
    HINSTANCE h;

    swprintf(py, MAX_PATH, L"%ls\\python\\python.exe", g_appdir);
    swprintf(script, MAX_PATH, L"%ls\\app\\serve_buffett_app.py", g_appdir);
    swprintf(args, 1200, L"-u \"%ls\" --no-browser --port %d", script, START_PORT);
    h = ShellExecuteW(NULL, L"open", py, args, g_appdir, SW_SHOWMINIMIZED);
    return (INT_PTR)h > 32;
}

/* ---------- 启动失败时：把 python 已输出的内容写入调试日志 ---------- */
static void dump_debug_log(void)
{
    char buf[8192];
    DWORD avail = 0, read = 0;
    HANDLE h;

    if (!g_hOutR) return;
    if (!PeekNamedPipe(g_hOutR, NULL, 0, NULL, &avail, NULL) || avail == 0) return;
    if (avail > sizeof buf - 1) avail = sizeof buf - 1;
    if (!ReadFile(g_hOutR, buf, avail, &read, NULL) || read == 0) return;
    buf[read] = 0;
    h = CreateFileW(g_debuglog, GENERIC_WRITE, 0, NULL,
                    CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (h != INVALID_HANDLE_VALUE) {
        DWORD w;
        WriteFile(h, buf, read, &w, NULL);
        CloseHandle(h);
    }
}

/* ---------- 定位 Edge ---------- */
static int find_edge(wchar_t *out, int cap)
{
    wchar_t base[MAX_PATH], p[MAX_PATH];
    const wchar_t *envs[] = { L"%ProgramFiles(x86)%", L"%ProgramFiles%" };
    int i;
    HKEY hk;
    DWORD sz;

    for (i = 0; i < 2; i++) {
        ExpandEnvironmentStringsW(envs[i], base, MAX_PATH);
        swprintf(p, MAX_PATH, L"%ls\\Microsoft\\Edge\\Application\\msedge.exe", base);
        if (GetFileAttributesW(p) != INVALID_FILE_ATTRIBUTES) {
            wcsncpy(out, p, cap - 1); out[cap - 1] = 0;
            return 1;
        }
    }
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE,
                      L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe",
                      0, KEY_READ, &hk) == ERROR_SUCCESS) {
        sz = sizeof p;
        if (RegQueryValueExW(hk, NULL, NULL, NULL, (BYTE *)p, &sz) == ERROR_SUCCESS &&
            GetFileAttributesW(p) != INVALID_FILE_ATTRIBUTES) {
            RegCloseKey(hk);
            wcsncpy(out, p, cap - 1); out[cap - 1] = 0;
            return 1;
        }
        RegCloseKey(hk);
    }
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
                      L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe",
                      0, KEY_READ, &hk) == ERROR_SUCCESS) {
        sz = sizeof p;
        if (RegQueryValueExW(hk, NULL, NULL, NULL, (BYTE *)p, &sz) == ERROR_SUCCESS &&
            GetFileAttributesW(p) != INVALID_FILE_ATTRIBUTES) {
            RegCloseKey(hk);
            wcsncpy(out, p, cap - 1); out[cap - 1] = 0;
            return 1;
        }
        RegCloseKey(hk);
    }
    return 0;
}

static void open_window(int port)
{
    /* 附加时间戳参数：单文件 HTML 更新后 Edge/WebView 可能命中旧版缓存
     *（表现为「已重新安装仍是旧数据」）。新 URL 必然回源。 */
    wchar_t url[96], edge[MAX_PATH], args[160];
    swprintf(url, 96, L"http://127.0.0.1:%d/?t=%lu", port, (unsigned long)time(NULL));
    if (find_edge(edge, MAX_PATH)) {
        swprintf(args, 160, L"--app=%ls --window-size=1280,840", url);
        ShellExecuteW(NULL, L"open", edge, args, NULL, SW_SHOWNORMAL);
    } else {
        ShellExecuteW(NULL, L"open", url, NULL, NULL, SW_SHOWNORMAL);
    }
}

/* 写「服务模式不可用」标记：本机安全软件持续拦截服务时，后续启动直接走离线模式 */
static void write_blocked_flag(void)
{
    HANDLE h = CreateFileW(g_flagfile, GENERIC_WRITE, 0, NULL,
                           CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (h != INVALID_HANDLE_VALUE) {
        DWORD w;
        WriteFile(h, "1", 1, &w, NULL);
        CloseHandle(h);
    }
}

/* 离线模式：直接打开内置 HTML（file://）——本地服务被安全软件拦截时
 * 也能完整使用阅读/搜索/划线/笔记（数据存浏览器 localStorage） */
static void open_file_mode(void)
{
    wchar_t path[500], url[1300], final_url[1400], edge[MAX_PATH], args[1450];
    int i, j;

    swprintf(path, 500, L"%ls\\app\\巴菲特投资智慧.html", g_appdir);
    wcscpy(url, L"file:///");
    j = 8;
    for (i = 0; path[i] && j < 1280; i++)
        url[j++] = (path[i] == L'\\') ? L'/' : path[i];
    url[j] = 0;
    /* 带时间戳，避免浏览器缓存旧版 HTML（离线模式同样需要） */
    swprintf(final_url, 1400, L"%ls?t=%lu", url, (unsigned long)time(NULL));

    if (find_edge(edge, MAX_PATH)) {
        swprintf(args, 1450, L"--app=%ls", final_url);
        ShellExecuteW(NULL, L"open", edge, args, NULL, SW_SHOWNORMAL);
    } else {
        ShellExecuteW(NULL, L"open", final_url, NULL, NULL, SW_SHOWNORMAL);
    }
    MessageBoxW(NULL,
                L"本地服务未能启动或已被安全软件终止，已改用离线模式打开。\n\n"
                L"阅读、搜索、划线与笔记功能正常；笔记/收藏保存在浏览器本地。\n\n"
                L"如需恢复完整功能（数据文件持久化）：\n"
                L"  1. 把安装目录加入杀毒软件信任区；\n"
                L"  2. 删除 %APPDATA%\\巴菲特投资智慧\\server-blocked.flag；\n"
                L"  3. 重新启动。",
                L"巴菲特投资智慧", MB_OK | MB_ICONINFORMATION);
}

/* 驻留监控：打开服务窗口后盯 10 秒，若服务被安全软件终止（探测失败），
 * 补开离线模式窗口并写标记，保证应用一定可用 */
static void monitor_and_fallback(int port)
{
    int t;
    for (t = 0; t < 10; t++) {
        Sleep(1000);
        if (!probe_port(port)) {
            write_blocked_flag();
            open_file_mode();
            return;
        }
    }
}

/* ---------- 入口 ---------- */
int wmain(void)
{
    WSADATA wsa;
    int port = 0;
    int blocked;

    WSAStartup(MAKEWORD(2, 2), &wsa);
    init_paths();
    blocked = GetFileAttributesW(g_flagfile) != INVALID_FILE_ATTRIBUTES;

    /* 1) 扫描完整回退端口段：服务已在运行则直接打开窗口。
     *    若上次服务被安全软件杀过（标记存在），不再尝试拉起，直接复用或离线 */
    port = find_existing_server();
    if (!port && !blocked) {
        Sleep(800);                        /* 消除「双开同时启动」竞态 */
        port = find_existing_server();
    }

    /* 2) 未运行 → 拉起内置 Python（三级兜底） */
    if (!port && !blocked) {
        if (!start_python(0))               /* a. 隐藏 */
            if (!start_python(1))           /* b. 可见最小化 */
                start_python_shell();       /* c. ShellExecute */
        port = wait_server();
    }

    /* 3) 服务就绪 → 打开应用窗口并驻留监控；否则转离线模式（文件直开） */
    if (port) {
        open_window(port);
        monitor_and_fallback(port);
    } else {
        write_blocked_flag();
        dump_debug_log();                  /* 留痕：%APPDATA%\巴菲特投资智慧\launch-debug.log */
        if (g_hProc) { WaitForSingleObject(g_hProc, 3000); CloseHandle(g_hProc); }
        if (g_hOutR) CloseHandle(g_hOutR);
        open_file_mode();
        return 3;
    }

    /* 4) 清理并返回 */
    if (g_hOutR) CloseHandle(g_hOutR);
    if (g_hProc) CloseHandle(g_hProc);
    return 0;
}
