# Excel to PDF Converter/Calculator

Only tested using vs-code.

A simple pdf to convert an Excel file to PDF with a policy number via a web interrface.

## Steps to setup locally.

1. Clone this repo:
   ```
   git clone https://github.com/Cleveerty/First-project.git
   cd First-project
   ```

2. Install dependencies:
   ```
   pip install
   flask
   openpyxl
   fpdf
   ```

3. Place your excel file into the project folder and rename it to `Book2.xlsx`. Or use the excel file provided and modifiy it how you see fit.

4. Run the app:
   ```
   python app.py
   ```

5. Open [http://127.0.0.1:5000/](http://127.0.0.1:5000/) in your browser or via the terminal.

## Files

- `app.py` — Flask backend
- `web.html` — Frontend form
- `requirements.txt` — Python dependencies
