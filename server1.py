import http.server
import socketserver
import os
import urllib.parse

PORT = 8000

class MyHttpRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed_path = urllib.parse.urlparse(self.path)
        path = parsed_path.path
        
        if path == '/':
            # 스크립트 파일이 있는 디렉토리로 이동
            script_dir = os.path.dirname(os.path.abspath(__file__))
            os.chdir(script_dir)
            
            # 현재 디렉토리에서 PDF 파일 목록을 가져옵니다.
            pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
            
            # 디버깅 메시지: 현재 디렉토리와 PDF 파일 목록 출력
            print(f"현재 디렉토리: {os.getcwd()}")
            print(f"PDF 파일 목록: {pdf_files}")
            
            # HTML 형식으로 PDF 파일 목록을 생성합니다.
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("<html><head><title>PDF Files</title></head><body>".encode('utf-8'))
            self.wfile.write("<h1>PDF Files</h1>".encode('utf-8'))
            if pdf_files:
                self.wfile.write("<ul>".encode('utf-8'))
                for pdf_file in pdf_files:
                    encoded_pdf_file = urllib.parse.quote(pdf_file)
                    self.wfile.write(f'<li><a href="/{encoded_pdf_file}">{pdf_file}</a></li>'.encode('utf-8'))
                self.wfile.write("</ul>".encode('utf-8'))
            else:
                self.wfile.write("<p>PDF 파일이 없습니다.</p>".encode('utf-8'))
            self.wfile.write("</body></html>".encode('utf-8'))
        else:
            # 요청된 경로를 URL 디코딩
            decoded_path = urllib.parse.unquote(path.strip('/'))
            
            # 스크립트 파일이 있는 디렉토리로 이동
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(script_dir, decoded_path)
            
            # 요청된 경로가 PDF 파일인지 확인
            if os.path.isfile(file_path) and file_path.endswith('.pdf'):
                # PDF 파일 제공
                self.send_response(200)
                self.send_header("Content-type", "application/pdf")
                self.end_headers()
                with open(file_path, 'rb') as file:
                    self.wfile.write(file.read())
            else:
                # 파일이 존재하지 않거나 PDF 파일이 아닌 경우 404 에러 반환
                self.send_error(404, f"파일을 찾을 수 없습니다: {self.path}")

# 웹 서버를 시작합니다.
with socketserver.TCPServer(("0.0.0.0", PORT), MyHttpRequestHandler) as httpd:
    print(f"포트 {PORT}에서 서버가 실행 중입니다.")
    httpd.serve_forever()