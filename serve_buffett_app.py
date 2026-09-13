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
import http.server
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
        if not self._prepare_static_path(path):
            self.send_error(404)
            return
        return super().do_GET()

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
