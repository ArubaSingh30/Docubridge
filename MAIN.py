import os

import cohere
import openai
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)
openai.api_key = os.environ["OPENAI_API_KEY"]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file     = request.files['excelFile']
    question = request.form['userQuestion']

    xls    = pd.ExcelFile(file, engine='openpyxl')
    sheets = xls.sheet_names
    df     = pd.read_excel(file, sheet_name=0, engine='openpyxl')
    df.dropna(how='all', inplace=True)
    table_html = df.head(10).to_html(classes='table table-striped',
                                     index=False)

    co = cohere.Client(os.environ["COHERE_API_KEY"])
    #message = cohere.ChatMessage(message="hello world!")
    response = co.chat(
        model="command-a-03-2025",
        message=question
    )
    print(response)
    ai_answer = response.text.strip()
    """"
    try:
        resp = openai.chat.completions.create(
          model="gpt-3.5-turbo",
          messages=[
            {"role":"system","content":"You are a financial-model assistant."},
            {"role":"user",  "content": "How are you"}
          ],
          max_tokens=150,
          temperature=0.7
        )
        if resp.choices[0].message.content is not None:
            ai_answer = resp.choices[0].message.content.strip()
        else:
            ai_answer = ""
    except openai.OpenAIError as e:
        ai_answer = f"OpenAI API error: {e}"
"""
    return render_template('result.html',
                           filename=file.filename,
                           question=question,
                           sheets=sheets,
                           table_html=table_html,
                           ai_answer=ai_answer)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
