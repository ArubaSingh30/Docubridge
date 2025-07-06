# AI Financial Analyst

A Flask web app that turns your Excel financial models into an interactive, AI-powered assistant.  
Upload a spreadsheet, ask questions in plain English, and get immediate insights backed by quick data profiling and Cohere’s Command-R model.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| Drag-and-drop upload | Supports `.xlsx`/`.xls` files up to **5 MB** |
| Data profiling | Detects Excel error tokens, missing values, calculates numeric stats, profit margin & basic time-series trends |
| Natural-language Q&A | Builds a rich prompt and queries Cohere for plain-English answers |
| Follow-up support | Ask additional questions without re-uploading; generates or explains Excel formulas on request |

---

## 🔧 Configuration

This project requires an OpenAI API key to interact with the GPT models. Follow these steps to set it up:

---

1. **Create a `.env` file**  
   In the root of the project (next to `main.py`), create a file named:
   ```bash
   touch .env

2. **Create&Activate an virtual environment**
   python -m venv venv
   source venv/bin/activate    # Windows: venv\Scripts\activate

3. **Install dependencies**
  pip install -r requirements.txt

### .env
FLASK_API_KEY=your_flask_secret_here
COHERE_API_KEY=8XXGiVGwbsWdEKq4zDIpcqWBQKg1IFMlbFHYf41G
OPENAI_API_KEY=--sk-your_real_openai_key_here--

---

### 💬 Example Questions

Try asking your AI Financial Analyst any of the following:

- “What were total revenues in 2023?”  
- “Calculate the month-over-month growth rate for Operating Expenses.”  
- “Which quarter had the highest net profit margin?”  
- “Show me a formula to compute Year-to-Date (YTD) cash flow in Excel.”  
- “Explain the formula in cell D15.”  
