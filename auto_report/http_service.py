import argparse
import hmac
import io
import json
import mimetypes
import os
import sys
import threading
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from config import Config
from main import run_report_job
from utils import setup_logging


class ReportHTTPServer(ThreadingHTTPServer):
    def __init__(
        self,
        server_address,
        RequestHandlerClass,
        logger,
        client_exe: str | None,
        save_dir: str | None,
        access_token: str,
    ):
        super().__init__(server_address, RequestHandlerClass)
        self.logger = logger
        self.client_exe = client_exe
        self.save_dir = save_dir
        self.access_token = access_token
        self.run_lock = threading.Lock()


class ReportHandler(BaseHTTPRequestHandler):
    server: ReportHTTPServer

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._write_json(200, {"status": "ok"})
            return
        if parsed.path == "/run":
            self._handle_run(parsed.query)
            return
        self._write_json(404, {"error": "not_found"})

    def _build_content_disposition(self, filename: str, fallback_name: str) -> str:
        safe_fallback = fallback_name.encode("ascii", errors="ignore").decode("ascii") or "download.bin"
        encoded = quote(filename, safe="")
        return f"attachment; filename=\"{safe_fallback}\"; filename*=UTF-8''{encoded}"

    def _handle_run(self, query_string: str) -> None:
        query = parse_qs(query_string)
        request_token = self._extract_request_token(query)
        if not hmac.compare_digest(request_token, self.server.access_token):
            self._write_json(401, {"error": "unauthorized"})
            return

        if not self.server.run_lock.acquire(blocking=False):
            self._write_json(409, {"error": "busy"})
            return

        client_exe = query.get("client_exe", [self.server.client_exe])[0]
        save_dir = query.get("save_dir", [self.server.save_dir])[0]
        try:
            code, results = run_report_job(
                client_exe=client_exe,
                save_dir=save_dir,
                launch_client=True,
                save_excel_when_no_launch=True,
                logger=self.server.logger,
            )
            if code != 0:
                self._write_json(500, {"error": "run_failed", "code": code})
                return
            if not results:
                self._write_json(500, {"error": "empty_report"})
                return
            if len(results) == 1:
                self._write_single_file(results[0])
                return
            self._write_zip(results)
        except Exception as exc:
            self.server.logger.error("HTTP任务执行失败: %s", exc)
            self._write_json(500, {"error": "internal_error", "message": str(exc)})
        finally:
            self.server.run_lock.release()

    def _write_single_file(self, file_path: Path) -> None:
        data = file_path.read_bytes()
        content_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
        fallback_name = f"report{file_path.suffix or '.bin'}"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", self._build_content_disposition(file_path.name, fallback_name))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _write_zip(self, files: list[Path]) -> None:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for file_path in files:
                zf.write(file_path, arcname=file_path.name)
        data = buffer.getvalue()
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", self._build_content_disposition("报表合集.zip", "reports.zip"))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _write_json(self, status_code: int, payload: dict) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _extract_request_token(self, query: dict[str, list[str]]) -> str:
        token_from_query = query.get("token", [""])[0]
        if token_from_query:
            return token_from_query

        token_from_header = self.headers.get("X-API-Token", "").strip()
        if token_from_header:
            return token_from_header

        authorization = self.headers.get("Authorization", "").strip()
        if authorization.lower().startswith("bearer "):
            return authorization[7:].strip()
        return ""

    def log_message(self, format, *args):
        self.server.logger.info("HTTP %s", format % args)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="auto_report_http_service")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--client-exe", default=None)
    parser.add_argument("--save-dir", default=None)
    parser.add_argument("--token", default=None)
    return parser.parse_args(argv)


def serve(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    config = Config()
    logger = setup_logging(config.log_dir)
    token = args.token or os.getenv("AUTO_REPORT_TOKEN") or config.http_token
    if not token:
        raise ValueError("HTTP token is required")
    server = ReportHTTPServer(
        (args.host, args.port),
        ReportHandler,
        logger=logger,
        client_exe=args.client_exe,
        save_dir=args.save_dir,
        access_token=token,
    )
    logger.info("HTTP服务已启动: http://%s:%s", args.host, args.port)
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(serve())
