from types import MethodDescriptorType
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    
    if 'excelFile' not in request.files:
        return 'No file part in the request', 400

    file = request.files['excelFile']
    question = request.form.get('userQuestion', '').strip()

    if file.filename == '':
        return 'No selected file', 400

    return render_template(
          'result.html',
          filename=file.filename,
          question=question
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
