from flask import Flask, send_from_directory, request
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
        <h1>파일 업로드</h1>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file"><br><br>
            <input type="submit" value="Upload">
        </form>
        <h1>파일 목록</h1>
        <ul>
            {}
        </ul>
    '''.format('<br>'.join(os.listdir(UPLOAD_FOLDER)))

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return '파일이 없습니다.'
    file = request.files['file']
    if file.filename == '':
        return '파일 이름이 없습니다.'

    file.save(os.path.join(UPLOAD_FOLDER, file.filename))
    return '파일이 업로드되었습니다!'

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)  # 구름 IDE에서 사용하는 포트로 설정