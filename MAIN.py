import os

import cohere
import openai
import pandas as pd
from flask import Flask, render_template, request
import json

app = Flask(__name__)
#openai.api_key = os.environ["OPENAI_API_KEY"]

@app.route('/')
def index():
    return render_template('index.html')
    
@app.route('/upload', methods=['POST'])
def upload():
    file     = request.files['excelFile']
    question = request.form['userQuestion']

    # Load first sheet of Excel
    df = pd.read_excel(file, sheet_name=0, engine='openpyxl')

    # 1) Detect Excel error tokens
    error_tokens = ['#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!']
    error_cells = []
    for r, row in df.iterrows():
        for col in df.columns:
            val = row[col]
            if isinstance(val, str) and any(token in val for token in error_tokens):
                cell = f"{str(col)}{str(r + 2)}"
                error_cells.append((cell, val))

    # 2) Detect missing values
    missing_cells = []
    for r, row in df.iterrows():
        for col in df.columns:
            if pd.isna(row[str(col)]):
                 missing_cells.append(f"{str(col)}{str(r + 2)}")

    # 3) Prepare HTML preview (show “ERROR” in place of tokens)
    df_clean  = df.replace(error_tokens, 'ERROR')
    table_html = df_clean.head(10).to_html(classes='table table-striped', index=False)

    # 4) Numeric stats
    num_rows     = len(df)
    cols         = df.columns.tolist()
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    stats_summary = []
    for col in numeric_cols:
        total = df[col].sum(skipna=True)
        mean  = df[col].mean(skipna=True)
        stats_summary.append(f"{col}: sum={total:.2f}, mean={mean:.2f}")
    stats_text = "; ".join(stats_summary) if stats_summary else "No numeric columns."

    # 4b) Detect time‐series trends if there’s a datetime column
    trend_summaries = []
    # Try to find any datetime column
    datetime_cols = df.select_dtypes(include=['datetime64', 'datetime64[ns]']).columns.tolist()
    if not datetime_cols:
        # as a fallback, see if the first column parses to datetime
        try:
            df = df.copy()
            df[cols[0]] = pd.to_datetime(df[cols[0]])
            datetime_cols = [cols[0]]
        except Exception:
            datetime_cols = []

    if datetime_cols:
        date_col = datetime_cols[0]
        df_ts = df.dropna(subset=[date_col]).sort_values(date_col)
        # for each numeric column, compute first vs last
        for col in numeric_cols:
            # only consider columns that align with the date index
            series = df_ts[[date_col, col]].dropna()
            if len(series) >= 2:
                first_val  = series.iloc[0][col]
                last_val   = series.iloc[-1][col]
                pct_change = (last_val - first_val) / first_val * 100 if first_val else None
                trend = ("increased" if last_val > first_val else
                         "decreased" if last_val < first_val else
                         "remained flat")
                if pct_change is not None:
                    trend_summaries.append(
                        f"{col} {trend} from {first_val:.2f} to {last_val:.2f} "
                        f"({pct_change:.1f}% over the period)"
                    )
        # optional: average period‐to‐period
        # you could compute series[col].pct_change().mean(), etc.

    trend_text = "; ".join(trend_summaries) if trend_summaries else "No clear time‐series trends detected."


    # 5) First 3 rows for context
    head_text = json.dumps(df.head(3).to_dict(orient='records'), indent=2)

    # 6) Build error note prefix
    notes = []
    if error_cells:
        entries = ", ".join(f"{c} ({v})" for c, v in error_cells)
        notes.append(f"Errors in cells: {entries}.")
    if missing_cells:
        sample = ", ".join(missing_cells[:5])
        suffix = "..." if len(missing_cells) > 5 else ""
        notes.append(f"Missing values in: {sample}{suffix}.")
    prefix = ("Note: " + " ".join(notes) + " ") if notes else ""

    # 7) Construct AI prompt
    prompt = (
        f"{prefix}Here is a spreadsheet with {num_rows} rows and columns {cols}. "
        f"Numeric column stats: {stats_text}. "
        f"Time‐series trends: {trend_text}. "
        f"First 3 rows: {head_text}. "
        f"The user asks: {question}. "
        f"Please answer based on the data provided above."
    )

    # 8) Call Cohere (or your LLM of choice)
    co = cohere.Client(os.environ['COHERE_API_KEY'])
    response = co.chat(model="command-a-03-2025", message=prompt)
    ai_answer = response.text.strip()

    return render_template(
        'result.html',
        filename=file.filename,
        question=question,
        table_html=table_html,
        ai_answer=ai_answer
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
