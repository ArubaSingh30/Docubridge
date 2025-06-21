import json
import os
import re
import uuid

import cohere
import markdown2
import pandas as pd
from flask import Flask, redirect, render_template, request, session, url_for
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ["FLASK_SECRET_KEY"]

UPLOAD_CACHE = {}

def process_dataframe(df, ws):
    """Run error/missing/stats/trend logic and return the pieces needed below."""
    # — error tokens
    error_tokens = ['#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!']
    error_cells = []
    for r, row in df.iterrows():
        for col in df.columns:
            v = row[col]
            if isinstance(v, str) and any(tok in v for tok in error_tokens):
                error_cells.append((f"{col}{r+2}", v))
    # — missing
    missing_cells = []
    for r, row in df.iterrows():
        for col in df.columns:
            if pd.isna(row[col]):
                missing_cells.append(f"{col}{r+2}")
    # — preview HTML
    df_clean   = df.replace(error_tokens, 'ERROR')
    table_html = df_clean.head(10).to_html(classes='table table-striped', index=False)
    # — numeric stats
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    stats = []
    for c in numeric_cols:
        total = df[c].sum(skipna=True)
        mean  = df[c].mean(skipna=True)
        stats.append(f"{c}: sum={total:.2f}, mean={mean:.2f}")
    stats_text = "; ".join(stats) if stats else "No numeric columns."
    # — trend
    datetime_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    trends = []
    if datetime_cols:
        dc = datetime_cols[0]
        df_ts = df.dropna(subset=[dc]).sort_values(dc)
        for c in numeric_cols:
            ser = df_ts[[dc, c]].dropna()
            if len(ser) >= 2:
                first, last = ser.iloc[0][c], ser.iloc[-1][c]
                pct = ((last - first) / first * 100) if first and pd.notnull(first) else None
                direction = ("increased" if last>first else "decreased" if last<first else "remained flat")
                if pct is not None:
                    trends.append(f"{c} {direction} from {first:.2f} to {last:.2f} ({pct:.1f}% over the period)")
    trend_text = "; ".join(trends) if trends else "No clear time-series trends detected."
    # — head text (first 3 rows JSON)
    head_text = json.dumps(df.head(3).to_dict(orient='records'), indent=2, default=str)
    # — prefix for errors/missing
    notes = []
    if error_cells:
        entries = ", ".join(f"{c} ({v})" for c, v in error_cells)
        notes.append(f"Errors in cells: {entries}.")
    if missing_cells:
        sample = ", ".join(missing_cells[:5])
        suffix = "..." if len(missing_cells) > 5 else ""
        notes.append(f"Missing values in: {sample}{suffix}.")
    prefix = ("Note: " + " ".join(notes) + " ") if notes else ""
    return table_html, prefix, stats_text, trend_text, head_text

def build_prompt(prefix, num_rows, cols, stats_text, trend_text, head_text, question):
    return (
        f"{prefix}Here is a spreadsheet with {num_rows} rows and columns {cols}. "
        f"Numeric column stats: {stats_text}. "
        f"Time-series trends: {trend_text}. "
        f"First 3 rows: {head_text}. "
        f"The user asks: {question}. "
        f"Please answer based on the data provided above."
    )
@app.route('/')
def index():
    return render_template('index.html')
    
@app.route('/upload', methods=['POST'])
def upload():
    file     = request.files['excelFile']
    question = request.form['userQuestion']

    df = pd.read_excel(file, sheet_name=0, engine='openpyxl', dtype=object)
    file.stream.seek(0)
    wb = load_workbook(file.stream, data_only=False)
    ws = wb.active
    table_html, prefix, stats_text, trend_text, head_text = process_dataframe(df, wb.active)
    prompt = build_prompt(prefix, len(df), df.columns.tolist(),
                          stats_text, trend_text, head_text, question)
    prompt = build_prompt(prefix, len(df), df.columns.tolist(),
          stats_text, trend_text, head_text, question) + "\n\nPlease format your answer in Markdown, using **bold** for emphasis."
    
    co   = cohere.Client(os.environ['COHERE_API_KEY'])
    resp = co.chat(model="command-a-03-2025", message=prompt)
    ai_answer = resp.text.strip()

    ai_html = markdown2.markdown(ai_answer)
    

    # Cache for follow-ups
    upload_id = str(uuid.uuid4())
    UPLOAD_CACHE[upload_id] = {'df': df, 'wb': wb, 'cols': df.columns.tolist()}
    session['upload_id'] = upload_id

    return render_template('assistant.html',
           table_html=table_html,
           ai_html=ai_html,
           filename=file.filename)

@app.route('/ask', methods=['POST'])
def ask():
    upload_id = session.get('upload_id')
    if not upload_id or upload_id not in UPLOAD_CACHE:
        return redirect(url_for('index'))

    data     = UPLOAD_CACHE[upload_id]
    df, wb, cols = data['df'], data['wb'], data['cols']
    question = request.form['userQuestion']

    # Re-calc preview & stats
    table_html, prefix, stats_text, trend_text, head_text = process_dataframe(df, wb.active)

    # Build prompt (or special formula case)
    m = re.search(r'cell\s*([A-Za-z]+[0-9]+)', question, re.IGNORECASE)
    if 'formula' in question.lower() and m:
        ref = m.group(1).upper()
        raw = wb.active[ref].value or ""
        prompt = f"Explain this Excel formula in simple terms:\n{raw}"
    else:
        prompt = build_prompt(prefix, len(df), cols, stats_text, trend_text, head_text, question)
        prompt += "\n\nPlease format your answer in Markdown, using **bold** for emphasis."

    # Call Cohere
    co   = cohere.Client(os.environ['COHERE_API_KEY'])
    resp = co.chat(model="command-a-03-2025", message=prompt)
    ai_answer = resp.text.strip()

    # Convert to HTML
    ai_html = markdown2.markdown(ai_answer)

    # Render with ai_html
    return render_template(
        'assistant.html',
        table_html=table_html,
        ai_html=ai_html,
        filename=None
    )
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
