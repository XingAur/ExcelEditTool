from __future__ import annotations

import socket
import threading
import time
import urllib.request

import uvicorn
import webview

from backend.app.main import create_app


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def run_server(port: int) -> None:
    config = uvicorn.Config(
        create_app(),
        host="127.0.0.1",
        port=port,
        log_level="warning",
        access_log=False,
    )
    server = uvicorn.Server(config)
    server.run()


def wait_for_server(port: int, timeout_seconds: float = 15) -> None:
    deadline = time.time() + timeout_seconds
    url = f"http://127.0.0.1:{port}/"
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise RuntimeError("客户端服务启动超时")


def main() -> None:
    port = find_free_port()
    thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    thread.start()
    wait_for_server(port)

    webview.create_window(
        "ExcelEditTool",
        f"http://127.0.0.1:{port}/",
        width=1280,
        height=820,
        min_size=(1120, 720),
    )
    webview.start()


if __name__ == "__main__":
    main()
