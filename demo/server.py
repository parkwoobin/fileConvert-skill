"""Loopback web demo for FileConvert's Windows desktop conversion engines."""
import argparse
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from shutil import copyfileobj
import subprocess
import sys
from tempfile import TemporaryDirectory
from threading import Lock
from urllib.parse import parse_qs, urlsplit

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))
from convert import PAIRS  # noqa: E402

MAX_UPLOAD = 32 * 1024 * 1024
LIMIT_SECONDS = 300
CONVERSION = Lock()
MODES = {
    (".pdf", ".docx"): {"objects": "--pdf-objects", "flow": None},
    (".hwpx", ".docx"): {"native": None, "structure": "--hwpx-structure"},
}
STATIC = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/style.css": ("style.css", "text/css; charset=utf-8"),
    "/app.js": ("app.js", "text/javascript; charset=utf-8"),
}


class Handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; object-src 'none'; frame-ancestors 'none'; base-uri 'none'")
        super().end_headers()

    def reply(self, status, data, kind):
        self.send_response(status)
        self.send_header("Content-Type", kind)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def json_reply(self, status, data):
        self.reply(status, json.dumps(data, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def local_request(self):
        port = self.server.server_port
        allowed = {f"127.0.0.1:{port}", f"localhost:{port}"}
        host = self.headers.get("Host")
        if host not in allowed:
            self.json_reply(403, {"message": "로컬 데모 주소로 접속해 주세요."})
            return False
        if self.command == "POST" and self.headers.get("Origin") != f"http://{host}":
            self.json_reply(403, {"message": "데모 페이지에서 변환을 시작해 주세요."})
            return False
        return True

    def do_GET(self):
        if not self.local_request():
            return
        path = urlsplit(self.path).path
        if path == "/api/formats":
            self.json_reply(200, {
                "pairs": sorted([a[1:], b[1:]] for a, b in PAIRS),
                "maxBytes": MAX_UPLOAD,
            })
        elif path in STATIC:
            name, kind = STATIC[path]
            self.reply(200, (HERE / name).read_bytes(), kind)
        else:
            self.json_reply(404, {"message": "페이지를 찾을 수 없습니다."})

    def do_POST(self):
        if not self.local_request():
            return
        route = urlsplit(self.path)
        if route.path != "/api/convert":
            self.json_reply(404, {"message": "변환 경로를 찾을 수 없습니다."})
            return
        options = parse_qs(route.query)
        source = "." + options.get("from", [""])[0]
        target = "." + options.get("to", [""])[0]
        selection = options.get("mode", ["native"])[0]
        modes = MODES.get((source, target), {"native": None})
        fidelity = options.get("fidelity", ["strict"])[0]
        if (source, target) not in PAIRS or selection not in modes or fidelity not in {"strict", "same-pages"}:
            self.json_reply(400, {"message": "지원하는 변환 방식과 출력 형식을 선택해 주세요."})
            return
        if self.headers.get_content_type() != "application/octet-stream":
            self.json_reply(415, {"message": "변환할 파일을 첨부해 주세요."})
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if not 0 < length <= MAX_UPLOAD:
            self.json_reply(413, {"message": "비어 있지 않은 32MB 이하 파일을 선택해 주세요."})
            return
        if not CONVERSION.acquire(blocking=False):
            self.json_reply(429, {"message": "다른 파일을 변환하고 있어요. 잠시 후 다시 시도해 주세요."})
            return
        try:
            self.run_conversion(length, source, target, modes[selection], fidelity)
        except (BrokenPipeError, ConnectionResetError, TimeoutError):
            self.close_connection = True
        finally:
            CONVERSION.release()

    def run_conversion(self, length, source, target, mode, fidelity):
        with TemporaryDirectory(prefix="getconvert-demo-") as folder:
            root = Path(folder)
            input_file = root / ("input" + source)
            output_file = root / ("converted" + target)
            report_file = root / "report.json"
            self.connection.settimeout(30)
            with input_file.open("wb") as uploaded:
                while length:
                    chunk = self.rfile.read(min(length, 65536))
                    if not chunk:
                        self.json_reply(400, {"message": "업로드가 중단되었습니다."})
                        return
                    uploaded.write(chunk)
                    length -= len(chunk)
            command = [sys.executable, str(SCRIPTS / "convert.py"),
                       str(input_file), str(output_file), "--report", str(report_file)]
            if mode:
                command.append(mode)
            if fidelity == "same-pages":
                command.append("--allow-differences")
            try:
                result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                        timeout=LIMIT_SECONDS, check=False)
            except subprocess.TimeoutExpired:
                self.json_reply(504, {"message": "변환 시간이 초과됐습니다. 파일을 나눠서 다시 시도해 주세요."})
                return
            summary = {}
            if report_file.exists():
                report = json.loads(report_file.read_text(encoding="utf-8"))
                summary = {key: report[key] for key in ("source_pages", "output_pages", "visual_match")}
            if result.returncode or not output_file.is_file():
                if result.returncode == 2 and summary:
                    if summary["source_pages"] != summary["output_pages"]:
                        code = "page-count-changed"
                        message = "페이지 수가 달라 결과 파일을 제공할 수 없습니다."
                    else:
                        code = "visual-difference"
                        message = "변환은 완료됐지만, 원본과 화면이 달라 현재 설정으로는 다운로드할 수 없습니다."
                else:
                    code = "conversion-failed"
                    message = "파일 형식과 암호 설정, 변환에 필요한 Office 앱이 설치되어 있는지 확인해 주세요."
                self.json_reply(422, {"code": code, "message": message, **summary})
                return
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="converted{target}"')
            self.send_header("Content-Length", str(output_file.stat().st_size))
            self.send_header("X-Conversion-Report", json.dumps(summary))
            self.end_headers()
            with output_file.open("rb") as result_file:
                copyfileobj(result_file, self.wfile)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    with ThreadingHTTPServer(("127.0.0.1", args.port), Handler) as server:
        print(f"겟컨버트 데모: http://127.0.0.1:{server.server_port}/", flush=True)
        server.serve_forever()


if __name__ == "__main__":
    main()
