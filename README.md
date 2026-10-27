# ECONOMIC EVALUATION GROUP 2 (FIELD DEVELOPMENT)

A comprehensive Python and Flask-based economic evaluation suite designed for petroleum field development projects. It performs cash flow modeling, fiscal terms analysis, depreciation modeling, sensitivity analysis, and generates interactive reports and downloadable packages.

## Features

- **Comprehensive Economic Engine**:
  - Discounted Cash Flow (DCF) modeling.
  - Calculation of key petroleum economic indicators: NPV, IRR, Payback Period, Profitability Index (PI / PIR), and Unit Technical Cost (UTC).
  - Multiple depreciation methods (SL, DB, DDB, UOP).
  - Royalties, production sharing/tax schemes, and economic limit cutoff options.
- **Sensitivity & Scenario Analysis**:
  - Capex, Opex, Oil Price, and Production sensitivities (-30% to +30%).
  - Spider plots and Tornado charts.
- **Interactive Web Interface**:
  - Flask backend with responsive UI.
  - Direct file upload for Excel datasets or one-click run with the sample template.
  - Dynamic figure preview and instant ZIP archive download containing full charts, sensitivity matrices, and summary reports.
- **Customizable Export & Saving**:
  - Save results to local preset folders or custom export paths.

## Project Structure

```
├── app.py                         # Flask web application & REST API
├── economic_engine.py             # Core petroleum economic calculations & figure generator
├── extracted_notebook_code.py     # Reference notebook extraction
├── oil field economics indicators updated.ipynb # Jupyter notebook evaluation model
├── sample_template.xlsx           # Sample dataset template for evaluation
├── static/
│   ├── app.css                    # Modern UI stylesheet
│   └── app.js                     # Frontend logic and charts viewer
├── templates/
│   └── index.html                 # Web dashboard template
├── test_app.py                    # Comprehensive test suite for Flask endpoints
├── test_engine.py                 # Smoke test for economic engine
├── requirements.txt               # Python package dependencies
└── .gitignore                     # Git ignore rules
```

## Quick Start

### 1. Prerequisites & Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/smart11524/EconomicIndicator.git
cd EconomicIndicator
pip install -r requirements.txt
```

### 2. Running the Application

Start the Flask server:

```bash
python app.py
```

Open your browser and navigate to:
```
http://127.0.0.1:5000
```

### 3. Running Tests

Execute the automated test suite:

```bash
python test_app.py
```

## License

MIT License
