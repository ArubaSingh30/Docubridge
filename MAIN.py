from types import MethodDescriptorType
from flask import Flask, render_template, request
import pandas
import os
import pandas as pd

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
    
    xls = pd.ExcelFile(file, engine='openpyxl')
    sheets = xls.sheet_names

    df = pd.read_excel(file, engine='openpyxl', sheet_name=0)
    df.dropna(how='all', inplace=True)

    #col2_sum = df.iloc[:, 1].sum()
    
    table_html = df.head(10).to_html(classes='table table-striped', index=False)
    
    return render_template(
        'result.html',
        filename=file.filename,
        question=question,
        sheets=sheets,
        #col2_sum=col2_sum,
        table_html=table_html
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
