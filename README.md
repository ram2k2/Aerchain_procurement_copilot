# Aerchain RFx Procurement Copilot

## Run on Windows

1. Install Python 3.10 or later.
2. From this folder, create and activate a project-local environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```

3. Set a Gemini API key in the same PowerShell session:

   ```powershell
   $env:GEMINI_API_KEY = "your-key"
   ```

4. Start Streamlit:

   ```powershell
   streamlit run app.py
   ```

The app starts with an empty RFx. Enter the buyer's scope, line items, questionnaire, and terms under **User Reference**, then generate and review the RFx. Upload up to five vendor responses in their original formats; vendors do not need to follow a template. At least one RFx line item is required before responses can be processed. `RFx.docx` and the vendor files in this folder are reference/sample inputs used for local parsing and offline tests; the app does not preload the reference RFx or show a sample-demo button. `Vendor_1.xlsx` is processed without Gemini; non-spreadsheet uploads are sent together in one extraction request.

## Gemini request budget

- RFx drafting: one request per explicit Generate click.
- Vendor processing: zero requests for spreadsheet-only uploads; one request for the entire non-spreadsheet batch.
- Comparison: local Python logic; no Gemini request.
- Analyst: one request for a new question and current comparison. Repeating the same question against unchanged data reuses the answer in the Streamlit session.

The app uses `gemini-3.5-flash-lite`, selected for low-cost multimodal document parsing and structured output.

Comparison uses the RFx generated in the current session. The Analyst receives normalized extracted values and source evidence, not the original documents. Currency conversion is not applied; cross-currency totals are excluded.

## Tests

Run the offline checks with:

```powershell
python -m unittest test_main.py
```

These tests do not use Gemini or consume API quota. A real extraction/Analyst run requires a valid Gemini key and available quota.
