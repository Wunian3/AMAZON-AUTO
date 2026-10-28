"""Local single-user web app. Uploads are processed in memory, never stored."""
import argparse
import base64
import binascii
import json
import logging
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from calculator import MAX_UPLOAD, calculate, export_workbook, inspect

STATIC = Path(__file__).parent / 'static'


class Handler(BaseHTTPRequestHandler):
    def json_response(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False, allow_nan=False).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == '/health':
            return self.json_response(200, {'status': 'ok'})
        routes = {'/': ('index.html', 'text/html; charset=utf-8'),
                  '/app.js': ('app.js', 'application/javascript; charset=utf-8'),
                  '/style.css': ('style.css', 'text/css; charset=utf-8')}
        if self.path not in routes:
            return self.json_response(404, {'error': '页面不存在'})
        name, mime = routes[self.path]
        body = (STATIC / name).read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'")
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path not in ('/api/inspect', '/api/calculate'):
            return self.json_response(404, {'error': '接口不存在'})
        try:
            # JSON-only same-origin API; no permissive CORS or form submissions.
            if self.headers.get('Content-Type', '').split(';')[0] != 'application/json':
                return self.json_response(415, {'error': '仅接受 JSON 请求'})
            origin = self.headers.get('Origin')
            if origin and origin not in ('http://' + self.headers.get('Host', ''),
                                         'https://' + self.headers.get('Host', '')):
                return self.json_response(403, {'error': '不接受跨站请求'})
            size = int(self.headers.get('Content-Length', '0'))
            if size <= 0 or size > MAX_UPLOAD * 4 // 3 + 20000:
                return self.json_response(413, {'error': '请求为空或文件超过20 MB'})
            payload = json.loads(self.rfile.read(size))
            if not isinstance(payload, dict) or not isinstance(payload.get('file'), str):
                raise ValueError('缺少工作簿文件')
            content = base64.b64decode(payload['file'], validate=True)
            if self.path == '/api/inspect':
                return self.json_response(200, inspect(content))
            options = payload.get('options', {})
            if not isinstance(options, dict):
                raise ValueError('参数格式错误')
            report = calculate(content, options)
            report['excel'] = base64.b64encode(export_workbook(report)).decode('ascii')
            self.json_response(200, report)
        except (ValueError, TypeError, binascii.Error) as exc:
            self.json_response(400, {'error': str(exc)})
        except Exception:
            logging.exception('Workbook processing failed')
            self.json_response(500, {'error': '工作簿处理失败，请检查文件格式或查看服务日志'})


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8000)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f'AMAZON-AUTO listening on {args.host}:{args.port}', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
