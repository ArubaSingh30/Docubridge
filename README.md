# AI Financial Analyst

Turn your Excel financial models into an interactive AI-powered assistant.
Upload your spreadsheet, ask plain-English questions, and get smart, data-backed answers instantly.

---

## Key Features

| Feature | Description |
|---------|-------------|
| Drag-and-drop upload | Accepts `.xlsx`/`.xls` files up to **5 MB** |
| Data profiling | Detects Excel error tokens, missing values, calculates numeric stats, profit margin & basic time-series trends |
| Natural-language Q&A | Builds a rich prompt and queries Cohere for plain-English answers |
| Follow-up support | Ask additional questions without re-uploading; generates or explains Excel formulas on request |

---

## Setup Instructions (On Replit)

This project requires an API key to interact with the AI models. Follow these steps to set it up:

---

1. **Fork the Replit project**
   Visit https://replit.com/@bacilasebi/DocuBridge-Financial-Model?v=1 and click Fork or Run to launch your own copy

2. **Install dependencies**
   Replit will usually auto-install from requirements.txt. If not, open the Shell and run:
   ```bash
   pip install -r requirements.txt

3. **Create a .env file**
   In Replit, go to the Secrets tab (🔐 icon in sidebar), and add the following variables:
   ```ini
   FLASK_API_KEY = your_flask_secret_here
   COHERE_API_KEY = your_cohere_key_here
   OPENAI_API_KEY = your_openai_key_here

---

## Setup Instructions (local)

This project requires an API key to interact with the GPT models. Follow these steps to set it up:

---

1. **Create a `.env` file**  
   In the root of the project (next to `main.py`), create a file named:
   ```bash
   touch .env

2. **Create&Activate an virtual environment**
   python -m venv venv
   source venv/bin/activate    # On Windows: venv\Scripts\activate

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt

4. **Add your API keys
   Create a .env file in the project root and add:
   ```bash
   FLASK_API_KEY=your_flask_secret_here
   COHERE_API_KEY=your_cohere_key_here
   OPENAI_API_KEY=your_openai_key_here

### Example Questions

Try asking your AI Financial Analyst any of the following:

- “What were total revenues in 2023?”  
- “Calculate the month-over-month growth rate for Operating Expenses.”  
- “Which quarter had the highest net profit margin?”  
- “Show me a formula to compute Year-to-Date (YTD) cash flow in Excel.”  
- “Explain the formula in cell D15.”

### Preview
