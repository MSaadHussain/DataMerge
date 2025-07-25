# 🧠 Product Matcher (Excel + Gemini AI + GUI)

A Python GUI tool that helps match product names between two Excel sheets using fuzzy matching and Google Gemini AI. The app suggests the best matching product and updates price fields accordingly.

---

## 📁 Project Structure

project-folder/
│
├── updated_benchmark.xlsx # Benchmark products (input/output)
├── store_prices.xlsx # Store products and prices
├── progress.json # Auto-generated progress tracker
├── matcher.py # Main script (this code)
└── README.md # This file


---

## 🧰 Requirements

- Python 3.9 or above
- Internet connection
- Google Gemini API Key

---

## ⚙️ Setup Instructions

### ✅ Step 1: Install Python

Download and install Python from:  
[https://www.python.org/downloads/](https://www.python.org/downloads/)

Make sure to check ✅ `Add Python to PATH` during setup.

Check installation:

```bash
python --version
```

##  📦 Step 2: Install Required Python Packages
Optionally create a virtual environment:

```bash
python -m venv venv
```
Activate it:

Windows:

bash
```
venv\Scripts\activate
macOS/Linux:
```
bash
```
source venv/bin/activate
Then install the dependencies:
```

bash
```
pip install openpyxl rapidfuzz google-generativeai
```
