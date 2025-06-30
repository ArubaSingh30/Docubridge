import cohere
import json
import logging
import markdown2
import os
import pandas as pd
import re
import uuid
from flask import (Flask, flash, redirect, render_template, request, session,
                   url_for)
from openpyxl import load_workbook
from werkzeug.exceptions import RequestEntityTooLarge
from zipfile import BadZipFile

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_API_KEY", "a-default-secret-key-for-development")

# --- Configuration ---
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 
ALLOWED_EXTENSIONS = {'.xls', '.xlsx'}

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

UPLOAD_CACHE = {}

def process_dataframe(df):
    """Analyzes a DataFrame to extract metadata, stats, and trends."""
    # --- Data Quality Checks ---
    error_tokens = ['#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!']
    error_cells = []
    for r_idx, row in df.iterrows():
        for c_idx, value in row.items():
            if isinstance(value, str) and any(tok in value for tok in error_tokens):
                error_cells.append((f"{c_idx}{r_idx + 2}", value))

    missing_cells = []
    for r_idx, row in df.iterrows():
        for c_idx, value in row.items():
            if pd.isna(value):
                missing_cells.append(f"{c_idx}{r_idx + 2}")

    # --- Data Preview ---
    df_clean = df.replace(error_tokens, 'ERROR')
    table_html = df_clean.head(10).to_html(classes='table table-striped', index=False, border=0)

    # --- Numeric Stats ---
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    stats = [f"{c}: sum={df[c].sum(skipna=True):,.2f}, mean={df[c].mean(skipna=True):,.2f}" for c in numeric_cols]
    stats_text = "; ".join(stats) if stats else "No numeric columns found."

    # --- FEATURE: Enhanced Trend Analysis & Ratio Calculation ---
    trends = []
    ratios = []
    datetime_cols = df.select_dtypes(include=['datetime64']).columns.tolist()

    # Simple check for common column names for ratio calculation
    revenue_col = next((c for c in df.columns if 'revenue' in c.lower()), None)
    profit_col = next((c for c in df.columns if 'profit' in c.lower() or 'income' in c.lower()), None)
    if revenue_col and profit_col and df[revenue_col].sum() > 0:
        profit_margin = (df[profit_col].sum() / df[revenue_col].sum()) * 100
        ratios.append(f"Overall Profit Margin: {profit_margin:.2f}%")

    if datetime_cols:
        dc = datetime_cols[0]
        df_ts = df.dropna(subset=[dc]).sort_values(by=dc)
        for c in numeric_cols:
            ser = df_ts[[dc, c]].dropna()
            if len(ser) >= 2:
                # Month-over-month trend
                ser[dc] = pd.to_datetime(ser[dc])
                monthly_change = ser.set_index(dc).resample('M')[c].sum().pct_change().mean() * 100
                if pd.notnull(monthly_change):
                    direction = "grew" if monthly_change > 0 else "declined"
                    trends.append(f"{c} {direction} by an average of {monthly_change:.1f}% month-over-month")

    trend_text = "; ".join(trends) if trends else "No clear time-series trends detected."
    ratio_text = "; ".join(ratios) if ratios else "No key financial ratios calculated."

    # --- Contextual Information ---
    head_text = json.dumps(df.head(3).to_dict(orient='records'), indent=2, default=str)
    notes = []
    if error_cells:
        notes.append(f"Errors found in cells like: {error_cells[0][0]}.")
    if missing_cells:
        notes.append(f"Missing values found in cells like: {missing_cells[0]}.")
    prefix = ("Note: " + " ".join(notes) + " ") if notes else ""

    return table_html, prefix, stats_text, trend_text, ratio_text, head_text

def build_prompt(prefix, num_rows, cols, stats_text, trend_text, ratio_text, head_text, question):
    """Builds the final prompt for the language model."""
    return (
        f"{prefix}Analyze the following spreadsheet data. "
        f"The sheet has {num_rows} rows and columns: {cols}. "
        f"Key Stats: {stats_text}. "
        f"Financial Ratios: {ratio_text}. "
        f"Time-series Trends: {trend_text}. "
        f"Data Preview (first 3 rows): {head_text}. "
        f"Given this context, please answer the user's question: '{question}'"
    )

@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    logging.warning("File upload failed: Exceeded size limit.")
    flash("File is too large. Please upload a file smaller than 5 MB.")
    return redirect(url_for('index'))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'excelFile' not in request.files or not request.files['excelFile'].filename:
        flash("No file selected. Please choose an Excel file.")
        return redirect(url_for('index'))

    file = request.files['excelFile']
    question = request.form['userQuestion']

    filename = file.filename
    file_ext = os.path.splitext(filename)[1].lower()
    if file_ext not in ALLOWED_EXTENSIONS:
        logging.warning(f"Invalid file type: {filename}")
        flash("Invalid file type. Please upload an Excel file (.xls, .xlsx).")
        return redirect(url_for('index'))

    logging.info(f"File uploaded: {filename}, Question: '{question}'")

    try:
        # FEATURE: Multi-sheet handling. We'll analyze the first sheet by default.
        # More advanced implementation could let the user choose a sheet.
        df = pd.read_excel(file, sheet_name=0, dtype=object)
    except (ValueError, BadZipFile) as e:
        logging.error(f"Corrupt Excel file: {filename}, error: {e}")
        flash("Sorry, I couldn’t read that file. Please check the format.")
        return redirect(url_for('index'))

    wb = None
    if file_ext == '.xlsx':
        try:
            file.stream.seek(0)
            wb = load_workbook(file.stream, data_only=False)
        except Exception:
            pass 

    table_html, prefix, stats, trends, ratios, head = process_dataframe(df)

    prompt = build_prompt(prefix, len(df), df.columns.tolist(), stats, trends, ratios, head, question)
    logging.info(f"Generated prompt (length: {len(prompt)})")

    try:
        co = cohere.Client(os.environ.get('COHERE_API_KEY'))
        resp = co.chat(model="command-r", message=prompt)
        ai_answer = resp.text.strip()
    except Exception as e:
        logging.error(f"Cohere API call failed: {e}")
        flash("The AI service is currently unavailable, please try again later.")
        return redirect(url_for('index'))

    ai_html = markdown2.markdown(ai_answer)

    upload_id = str(uuid.uuid4())
    UPLOAD_CACHE[upload_id] = {'df': df, 'wb': wb, 'cols': df.columns.tolist()}
    session['upload_id'] = upload_id

    return render_template('assistant.html',
                           result={'question': question, 'answer': ai_html},
                           table_html=table_html,
                           filename=filename)

@app.route('/ask', methods=['POST'])
def ask():
    upload_id = session.get('upload_id')
    if not upload_id or upload_id not in UPLOAD_CACHE:
        flash("Your session has expired. Please upload a file again.")
        return redirect(url_for('index'))

    data = UPLOAD_CACHE[upload_id]
    df, wb, cols = data['df'], data['wb'], data['cols']
    question = request.form['userQuestion']
    logging.info(f"Follow-up question for {upload_id}: '{question}'")

    table_html, prefix, stats, trends, ratios, head = process_dataframe(df)

    # --- FEATURE: Excel Formula Generation ---
    formula_keywords = ['how do i calculate', 'what is the formula for', 'excel formula for']
    if any(keyword in question.lower() for keyword in formula_keywords):
        prompt = f"Please provide an Excel formula to answer the following user question. Be concise and provide a brief explanation. Question: {question}"
    else:
        m = re.search(r'cell\s*([A-Za-z]+[0-9]+)', question, re.IGNORECASE)
        if wb and 'formula' in question.lower() and m:
            ref = m.group(1).upper()
            try:
                raw = wb.active[ref].value or ""
                prompt = f"Explain this Excel formula in simple terms: {raw}" if raw else f"Cell {ref} is empty."
            except Exception:
                prompt = f"Could not access cell {ref}."
        else:
            prompt = build_prompt(prefix, len(df), cols, stats, trends, ratios, head, question)

    logging.info(f"Generated follow-up prompt (length: {len(prompt)})")

    try:
        co = cohere.Client(os.environ.get('COHERE_API_KEY'))
        resp = co.chat(model="command-r", message=prompt)
        ai_answer = resp.text.strip()
    except Exception as e:
        logging.error(f"Cohere API call failed on follow-up: {e}")
        return "The AI service is currently unavailable.", 503

    ai_html = markdown2.markdown(ai_answer)

    return render_template('assistant.html',
                           result={'question': question, 'answer': ai_html},
                           table_html=table_html,
                           filename=None)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)