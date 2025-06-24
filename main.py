import json
import logging
import os
import re
import uuid

import cohere
import markdown2
import pandas as pd
from flask import (Flask, flash, redirect, render_template, request, session,
                   url_for)
from openpyxl import load_workbook
from werkzeug.exceptions import RequestEntityTooLarge
from zipfile import BadZipFile

app = Flask(__name__)
# It's recommended to use a more secure, permanent secret key for production apps
app.secret_key = os.environ.get("FLASK_API_KEY", "a-default-secret-key-for-development")

# --- Configuration ---
# Set a file size limit (e.g., 5 MB) to prevent large uploads
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 
ALLOWED_EXTENSIONS = {'.xls', '.xlsx'}

# --- Logging Setup ---
# Configure basic logging to print to the console for debugging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# In-memory cache for storing uploaded file data.
# Note: This is temporary and will be cleared on server restart.
UPLOAD_CACHE = {}

def process_dataframe(df):
    """Analyzes a DataFrame to extract metadata, stats, and trends."""
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

    df_clean = df.replace(error_tokens, 'ERROR')
    table_html = df_clean.head(10).to_html(classes='table table-striped', index=False, border=0)

    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    stats = []
    for c in numeric_cols:
        total = df[c].sum(skipna=True)
        mean = df[c].mean(skipna=True)
        stats.append(f"{c}: sum={total:,.2f}, mean={mean:,.2f}")
    stats_text = "; ".join(stats) if stats else "No numeric columns found."

    datetime_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    trends = []
    if datetime_cols:
        dc = datetime_cols[0]
        df_ts = df.dropna(subset=[dc]).sort_values(by=dc)
        for c in numeric_cols:
            ser = df_ts[[dc, c]].dropna()
            if len(ser) >= 2:
                first, last = ser[c].iloc[0], ser[c].iloc[-1]
                if pd.notnull(first) and first != 0:
                    pct = ((last - first) / first) * 100
                    direction = "increased" if last > first else "decreased"
                    trends.append(f"{c} {direction} from {first:,.2f} to {last:,.2f} ({pct:.1f}% over the period)")
    trend_text = "; ".join(trends) if trends else "No clear time-series trends detected."

    head_text = json.dumps(df.head(3).to_dict(orient='records'), indent=2, default=str)

    notes = []
    if error_cells:
        entries = ", ".join(f"{c} ({v})" for c, v in error_cells[:5])
        notes.append(f"Errors found in cells: {entries}{'...' if len(error_cells) > 5 else ''}.")
    if missing_cells:
        sample = ", ".join(missing_cells[:5])
        notes.append(f"Missing values found in: {sample}{'...' if len(missing_cells) > 5 else ''}.")
    prefix = ("Note: " + " ".join(notes) + " ") if notes else ""

    return table_html, prefix, stats_text, trend_text, head_text

def build_prompt(prefix, num_rows, cols, stats_text, trend_text, head_text, question):
    """Builds the final prompt for the language model."""
    return (
        f"{prefix}Here is a summary of a spreadsheet with {num_rows} rows and columns named {cols}. "
        f"Key numeric stats: {stats_text}. "
        f"Time-series trends: {trend_text}. "
        f"Here are the first 3 rows as a sample: {head_text}. "
        f"Based on this information, the user asks: {question}. "
        f"Please provide a clear and concise answer based on the data provided."
    )

@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    """Handles file size errors."""
    logging.warning("File upload failed: File exceeded the size limit.")
    flash("File is too large. Please upload a file smaller than 5 MB.")
    return redirect(url_for('index'))

@app.route('/')
def index():
    """Renders the main upload page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    """Handles file upload, data processing, and the initial question."""
    if 'excelFile' not in request.files or not request.files['excelFile'].filename:
        flash("No file selected. Please choose an Excel file.")
        return redirect(url_for('index'))

    file = request.files['excelFile']
    question = request.form['userQuestion']

    # --- File Type Validation ---
    filename = file.filename
    file_ext = os.path.splitext(filename)[1].lower()
    if file_ext not in ALLOWED_EXTENSIONS:
        logging.warning(f"Invalid file type uploaded: {filename}")
        flash("Invalid file type. Please upload an Excel file (.xls, .xlsx).")
        return redirect(url_for('index'))

    logging.info(f"File uploaded: {filename}, Question: '{question}'")

    try:
        df = pd.read_excel(file, sheet_name=0, dtype=object)
    except (ValueError, BadZipFile):
        logging.error(f"Failed to read corrupt Excel file: {filename}")
        flash("Sorry, I couldn’t read that file. Please check that it is not corrupted.")
        return redirect(url_for('index'))
    except Exception as e:
        logging.error(f"An unexpected error occurred while reading {filename}: {e}")
        flash(f"An unexpected error occurred while processing the file.")
        return redirect(url_for('index'))

    wb = None
    try:
        file.stream.seek(0)
        wb = load_workbook(file.stream, data_only=False)
    except Exception:
        pass # Expected for .xls files

    table_html, prefix, stats_text, trend_text, head_text = process_dataframe(df)

    prompt = build_prompt(prefix, len(df), df.columns.tolist(),
                          stats_text, trend_text, head_text, question)
    prompt += "\n\nPlease format your answer in Markdown, using **bold** for emphasis."
    logging.info(f"Generated prompt for Cohere API (length: {len(prompt)})")

    # --- AI API Error Handling ---
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
                           filename=file.filename)

@app.route('/ask', methods=['POST'])
def ask():
    """Handles follow-up questions."""
    upload_id = session.get('upload_id')
    if not upload_id or upload_id not in UPLOAD_CACHE:
        flash("Your session has expired. Please upload a file again.")
        return redirect(url_for('index'))

    data = UPLOAD_CACHE[upload_id]
    df, wb, cols = data['df'], data['wb'], data['cols']
    question = request.form['userQuestion']
    logging.info(f"Follow-up question for session {upload_id}: '{question}'")

    table_html, prefix, stats_text, trend_text, head_text = process_dataframe(df)

    m = re.search(r'cell\s*([A-Za-z]+[0-9]+)', question, re.IGNORECASE)

    if wb and 'formula' in question.lower() and m:
        ref = m.group(1).upper()
        try:
            raw = wb.active[ref].value or ""
            prompt = f"Explain this Excel formula in simple terms: {raw}" if raw else f"Cell {ref} is empty."
        except Exception:
            prompt = f"Could not access cell {ref}."
    else:
        prompt = build_prompt(prefix, len(df), cols, stats_text, trend_text, head_text, question)
        prompt += "\n\nPlease format your answer in Markdown, using **bold** for emphasis."
        if 'formula' in question.lower() and not wb:
            prompt += "\n\n*(Note: Formula inspection is only available for .xlsx files.)*"

    logging.info(f"Generated follow-up prompt (length: {len(prompt)})")

    # --- AI API Error Handling ---
    try:
        co = cohere.Client(os.environ.get('COHERE_API_KEY'))
        resp = co.chat(model="command-r", message=prompt)
        ai_answer = resp.text.strip()
    except Exception as e:
        logging.error(f"Cohere API call failed during follow-up: {e}")
        # For a follow-up, we might return an error without a full page redirect
        return "The AI service is currently unavailable, please try again later.", 503

    ai_html = markdown2.markdown(ai_answer)

    return render_template('assistant.html',
                           result={'question': question, 'answer': ai_html},
                           table_html=table_html,
                           filename=None)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
