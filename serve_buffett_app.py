#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巴菲特投资智慧.html 本地启动器
==============================

作用：
  1. 以本地 HTTP 服务（仅 127.0.0.1）托管「巴菲特投资智慧.html」；
  2. 通过同源 /api/llm-config 提供运行时 LLM 配置：API Key 从环境变量
     DEEPSEEK_API_KEY 读取（未设置时回退项目根 .env）。服务只把密钥保留在
     进程和页面内存中，不会将它写入状态或生成文件；
  3. 自动打开浏览器。

用法：
  python3 serve_buffett_app.py [--port 8666]
  或双击同目录的「启动巴菲特知识库.command」

环境变量（可选）：
  DEEPSEEK_API_KEY    API Key（缺省读项目根 .env）
  DEEPSEEK_API_BASE   API Base（缺省 https://api.deepseek.com/v1）
  BUFFETT_LLM_MODEL   模型（缺省 deepseek-v4-flash）

直接双击 html（file:// 方式）也可以使用，此时密钥为空，
可在应用「设置」面板手动填写。
"""

import argparse
import gzip
import html as html_mod
import http.server
import ipaddress
import re
import socket
import urllib.parse
import urllib.request
import zlib
from urllib.parse import quote, unquote
import json
import os
import socketserver
import sys
import tempfile
import threading
import webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = "巴菲特投资智慧.html"
DEFAULT_BASE = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-v4-flash"

# 记忆材料（笔记/收藏/已读/AI 对话）的本地持久化目录（用户级目录，天然不进 Git）：
#   macOS / Linux: ~/Library/Application Support/巴菲特投资智慧/
#   Windows:       %APPDATA%\巴菲特投资智慧\
#   可用环境变量 BUFFETT_DATA_DIR 覆盖（例如指向项目内时请在 .gitignore 忽略）
if os.name == "nt":
    _appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    DATA_DIR = os.environ.get("BUFFETT_DATA_DIR") or os.path.join(_appdata, "巴菲特投资智慧")
else:
    DATA_DIR = (os.environ.get("BUFFETT_DATA_DIR") or
                os.path.join(os.path.expanduser("~"), "Library", "Application Support", "巴菲特投资智慧"))
STATE_FILE = os.path.join(DATA_DIR, "state.json")


def load_state():
    """读取记忆材料 state.json；不存在或损坏时返回空 dict。"""
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            obj = json.load(f)
            return obj if isinstance(obj, dict) else {}
    except (OSError, ValueError):
        return {}


def save_state(obj):
    """原子写 state.json；每次使用唯一临时文件以支持并发请求/进程。"""
    os.makedirs(DATA_DIR, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".state-", suffix=".tmp", dir=DATA_DIR)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, STATE_FILE)
    finally:
        try:
            os.unlink(tmp)
        except FileNotFoundError:
            pass


def load_root_env():
    """读取项目根（服务脚本同目录）的 .env；找不到时返回 {}。"""
    p = os.path.join(HERE, ".env")
    if not os.path.isfile(p):
        return {}
    env = {}
    try:
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    env[k.strip()] = v.strip().strip('"').strip("'")
    except OSError:
        return {}
    return env


def resolve_llm_config():
    env = load_root_env()
    key = (os.environ.get("DEEPSEEK_API_KEY") or env.get("DEEPSEEK_API_KEY") or "").strip()
    base = (os.environ.get("DEEPSEEK_API_BASE") or env.get("DEEPSEEK_API_BASE") or DEFAULT_BASE).strip()
    model = (os.environ.get("BUFFETT_LLM_MODEL") or DEFAULT_MODEL).strip()
    return {"base": base, "key": key, "model": model}


# ===================== 联网检索 / 网页抓取（问答 AI 的研究工具） =====================
# 密钥来源：客户端设置面板填写（随请求头传入，不落盘）→ 环境变量 → 项目根 .env
SEARCH_PROVIDERS = {
    "bocha": {"env": "BOCHA_API_KEY", "label": "博查"},
    "tavily": {"env": "TAVILY_API_KEY", "label": "Tavily"},
}
FETCH_MAX_BYTES = 400_000
FETCH_MAX_CHARS = 20000
FETCH_TIMEOUT = 15
# 部分环境（企业代理 / 沙箱）会把公网域名解析到内网地址，导致抓取被 SSRF 防护拦下。
# 此时可设 BUFFETT_FETCH_ALLOW_PRIVATE=1 显式放开（默认保持严格防护）。
FETCH_ALLOW_PRIVATE = (os.environ.get("BUFFETT_FETCH_ALLOW_PRIVATE", "").strip().lower()
                       in ("1", "true", "yes"))
try:                       # 可选依赖：装了 brotli 才能解 br 压缩响应
    import brotli as _brotli
except ImportError:
    _brotli = None


def resolve_search_key(client_key="", provider=""):
    """返回 (provider, key)：客户端密钥优先，其次该 provider 的环境变量/.env。"""
    client_key = (client_key or "").strip()
    provider = (provider or "").strip().lower()
    env = load_root_env()
    if provider in SEARCH_PROVIDERS:
        names = [provider]
    else:
        names = ["bocha", "tavily"]
    for name in names:
        env_key = (os.environ.get(SEARCH_PROVIDERS[name]["env"])
                   or env.get(SEARCH_PROVIDERS[name]["env"]) or "").strip()
        if env_key:
            return name, (client_key or env_key)
    if client_key:
        return (provider if provider in SEARCH_PROVIDERS else "bocha"), client_key
    return "", ""


def _post_json(url, payload, headers, timeout=FETCH_TIMEOUT):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", errors="replace"))


def web_search(query, count=5, key="", provider="bocha"):
    """统一结构的联网检索：返回 {results:[{title,url,snippet}], message}。"""
    query = (query or "").strip()
    if not query:
        return {"results": [], "message": "缺少检索词"}
    try:
        count = max(1, min(int(count or 5), 8))
    except (TypeError, ValueError):
        count = 5
    if not key:
        return {"results": [], "message":
                "未配置联网检索密钥：可在应用「设置」面板填写搜索 API Key（博查/Tavily），"
                "或设置环境变量 BOCHA_API_KEY / TAVILY_API_KEY 后重启本地服务。"}
    try:
        if provider == "tavily":
            j = _post_json("https://api.tavily.com/search",
                           {"api_key": key, "query": query, "max_results": count,
                            "search_depth": "basic"},
                           {"Content-Type": "application/json"})
            items = [{"title": x.get("title", ""), "url": x.get("url", ""),
                      "snippet": (x.get("content") or "")[:400]}
                     for x in (j.get("results") or [])]
        else:
            j = _post_json("https://api.bochaai.com/v1/web-search",
                           {"query": query, "summary": True, "count": count},
                           {"Content-Type": "application/json",
                            "Authorization": "Bearer " + key})
            pages = (((j.get("data") or {}).get("webPages") or {}).get("value")) or []
            items = [{"title": x.get("name", ""), "url": x.get("url", ""),
                      "snippet": (x.get("summary") or x.get("snippet") or "")[:400]}
                     for x in pages]
        if not items:
            return {"results": [], "message": "联网检索无结果（可换关键词重试）"}
        return {"results": items[:count], "message": ""}
    except Exception as e:                                    # noqa: BLE001
        return {"results": [], "message": "联网检索失败：%s" % e}


def _host_is_public(host):
    """拒绝 loopback/内网/保留地址，避免本地服务被当作 SSRF 跳板。"""
    try:
        infos = socket.getaddrinfo(host, None)
    except OSError:
        return False
    for info in infos:
        try:
            addr = ipaddress.ip_address(info[4][0])
        except ValueError:
            return False
        if (addr.is_private or addr.is_loopback or addr.is_link_local
                or addr.is_reserved or addr.is_multicast or addr.is_unspecified):
            return False
    return True


def fetch_page(url):
    """抓取网页并抽取正文：返回 {title, text}。仅允许公网 http/https。"""
    parsed = urllib.parse.urlparse(url or "")
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        return {"title": "", "text": "", "message": "仅支持 http/https 网址"}
    if not FETCH_ALLOW_PRIVATE and not _host_is_public(parsed.hostname):
        return {"title": "", "text": "", "message":
                "拒绝访问内网或本机地址（如确认目标为公网、且本机使用代理，"
                "可设 BUFFETT_FETCH_ALLOW_PRIVATE=1 后重启本地服务）"}
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (compatible; BuffettWisdom/1.0; +local-research)",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate",
        })
        with urllib.request.urlopen(req, timeout=FETCH_TIMEOUT) as resp:
            raw = resp.read(FETCH_MAX_BYTES)
            charset = resp.headers.get_content_charset() or "utf-8"
            encoding = (resp.headers.get("Content-Encoding") or "").lower()
    except Exception as e:                                    # noqa: BLE001
        return {"title": "", "text": "", "message": "抓取失败：%s" % e}
    # 解压：按响应头 / 魔数处理 gzip、deflate；brotli 仅在已安装第三方库时支持
    if "br" in encoding and raw[:2] != b"\x1f\x8b":
        if _brotli:
            try:
                raw = _brotli.decompress(raw)
            except Exception:                                  # noqa: BLE001
                return {"title": "", "text": "", "message": "该站点返回 Brotli 压缩内容，解压失败"}
        else:
            return {"title": "", "text": "", "message":
                    "该站点返回 Brotli（br）压缩内容，本机 Python 无法解压；"
                    "可运行 python3 -m pip install brotli 后重启服务，或改用 web_search 的摘要/其他来源"}
    try:
        if "gzip" in encoding or raw[:2] == b"\x1f\x8b":
            raw = gzip.decompress(raw)
        elif "deflate" in encoding:
            raw = zlib.decompress(raw, -zlib.MAX_WBITS)
    except Exception:                                          # noqa: BLE001
        pass
    if raw[:5] == b"%PDF-":
        return {"title": "", "text": "", "message":
                "该网址是 PDF 文件，暂不支持直接解析正文；可改用 web_search 摘要，"
                "或查阅本地知识库中对应的信件/文章"}
    page = raw.decode(charset, errors="replace")
    title = ""
    m = re.search(r"<title[^>]*>(.*?)</title>", page, re.S | re.I)
    if m:
        title = re.sub(r"\s+", " ", html_mod.unescape(m.group(1))).strip()[:200]
    text = re.sub(r"(?is)<(script|style|noscript|svg|head)[^>]*>.*?</\1>", " ", page)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", html_mod.unescape(text)).strip()
    if not text:
        return {"title": title, "text": "", "message": "页面正文为空（可能需要登录或被反爬拦截）"}
    # 乱码检测：解码失败率过高时不把垃圾喂给模型
    if text.count("\ufffd") > max(20, len(text) * 0.05):
        return {"title": title, "text": "", "message":
                "页面内容无法正确解码（可能是压缩或编码异常），请改用其他来源"}
    return {"title": title, "text": text[:FETCH_MAX_CHARS], "message": ""}


class Handler(http.server.SimpleHTTPRequestHandler):
    """应用静态入口 + 同源配置/状态 API。"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=HERE, **kwargs)

    def end_headers(self):
        """所有响应一律禁止缓存。

        单文件应用更新后，若旧版 HTML 被 WebView/浏览器按启发式规则缓存，
        会出现「已重新安装但页面仍是旧数据」的现象（本地服务原先只对 API
        设置 no-store，静态 HTML 未设）。这里统一加上禁缓存头。
        """
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def _host_allowed(self):
        """只接受直接访问 loopback 的请求，避免 DNS rebinding。"""
        host = (self.headers.get("Host") or "").strip().lower()
        name = host.rsplit(":", 1)[0] if ":" in host else host
        return name in {"127.0.0.1", "localhost"}

    def _reject_bad_host(self):
        if self._host_allowed():
            return False
        self.send_error(403, "loopback host required")
        return True

    def _send_json(self, obj, extra_headers=None):
        body = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for key, value in (extra_headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def _prepare_static_path(self, path):
        """只暴露应用入口与无密钥的静态配置，拒绝仓库内其他文件。"""
        decoded = unquote(path)
        if decoded == "/":
            self.path = "/" + quote(INDEX, safe="")
            return True
        if decoded == "/" + INDEX:
            self.path = "/" + quote(INDEX, safe="")
            return True
        if decoded == "/llm-config.js":
            self.path = "/llm-config.js"
            return True
        return False

    def do_GET(self):
        if self._reject_bad_host():
            return
        path = self.path.split("?", 1)[0]
        if path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return
        if path == "/api/state":
            self._send_json(load_state())
            return
        if path == "/api/llm-config":
            cfg = resolve_llm_config()
            cfg["app"] = "buffett-wisdom"
            self._send_json(cfg, {"Cross-Origin-Resource-Policy": "same-origin"})
            return
        if path == "/api/websearch":
            self._handle_websearch()
            return
        if path == "/api/fetch":
            self._handle_fetch()
            return
        if not self._prepare_static_path(path):
            self.send_error(404)
            return
        return super().do_GET()

    # ---------- 问答 AI 的研究工具：联网检索 / 网页抓取 ----------
    def _query_params(self):
        qs = urllib.parse.urlsplit(self.path).query
        return {k: v[0] for k, v in urllib.parse.parse_qs(qs).items()}

    def _handle_websearch(self):
        params = self._query_params()
        provider, key = resolve_search_key(self.headers.get("X-Search-Api-Key", ""),
                                           self.headers.get("X-Search-Provider", ""))
        out = web_search(params.get("q", ""), params.get("k", 5), key, provider or "bocha")
        out["provider"] = SEARCH_PROVIDERS.get(provider, {}).get("label", "")
        self._send_json(out)

    def _handle_fetch(self):
        params = self._query_params()
        self._send_json(fetch_page(params.get("url", "")))

    def do_HEAD(self):
        if self._reject_bad_host():
            return
        path = self.path.split("?", 1)[0]
        if path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return
        if not self._prepare_static_path(path):
            self.send_error(404)
            return
        return super().do_HEAD()

    def do_PUT(self):
        """PUT /api/state：浏览器把记忆材料（笔记/收藏/已读/AI 对话）写回本地文件。

        整体替换语义（前端总是提交完整 6 键状态）；只接受白名单字段，
        绝不接收 settings（其中可能含 API Key）。服务端注入的密钥只存在于
        内存，用户在设置中手动填写的密钥只存在于浏览器 localStorage。
        """
        if self._reject_bad_host():
            return
        path = self.path.split("?", 1)[0]
        if path != "/api/state":
            self.send_response(404)
            self.end_headers()
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length else b"{}"
            obj = json.loads(raw.decode("utf-8"))
            if not isinstance(obj, dict):
                raise ValueError("state must be an object")
            allowed = {"notes", "favs", "read", "chat", "buffett_chat", "munger_chat"}
            clean = {k: v for k, v in obj.items() if k in allowed and isinstance(v, (dict, list))}
            save_state(clean)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        except (ValueError, OSError) as e:
            self.send_response(400)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(('{"ok":false,"error":%s}' % json.dumps(str(e))).encode("utf-8"))

    def log_message(self, fmt, *args):
        sys.stderr.write("[server] %s\n" % (fmt % args))


class ReuseTCPServer(socketserver.ThreadingTCPServer):
    """允许端口复用：开发/重启频繁时避免 TIME_WAIT 导致绑定失败。"""
    allow_reuse_address = True


def pick_port(preferred):
    for port in range(preferred, preferred + 20):
        try:
            srv = ReuseTCPServer(("127.0.0.1", port), Handler)
            return srv, port
        except OSError:
            continue
    raise SystemExit("[error] 端口 %d-%d 均被占用" % (preferred, preferred + 19))


def main():
    ap = argparse.ArgumentParser(description="巴菲特投资智慧 本地启动器")
    ap.add_argument("--port", type=int, default=int(os.environ.get("BUFFETT_PORT", "8666")))
    ap.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    args = ap.parse_args()

    cfg = resolve_llm_config()
    if cfg["key"]:
        print("[ok] LLM 密钥已从环境变量注入（DEEPSEEK_API_KEY，%d 位）" % len(cfg["key"]))
    else:
        print("[warn] 未找到 DEEPSEEK_API_KEY（环境变量 / 项目根 .env），"
              "可在应用「设置」面板手动填写密钥")

    srv, port = pick_port(args.port)
    url = "http://127.0.0.1:%d/%s" % (port, quote(INDEX, safe=''))  # 中文路径百分号编码，兼容任意默认浏览器
    print("[ok] 服务已启动: %s" % url)
    print("     模型: %s | 按 Ctrl+C 停止" % cfg["model"])
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n[bye] 已停止")
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
