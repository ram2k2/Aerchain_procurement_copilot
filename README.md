# Aerchain RFx Procurement Copilot

Live demo: [aerchain-procurement-copilot.streamlit.app](https://aerchain-procurement-copilot.streamlit.app/)

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
- Vendor processing: zero requests for spreadsheet-only uploads; one request for the entire non-spreadsheet batch, with up to 65,536 output tokens. If Gemini omits a file or returns invalid/incomplete JSON, processing stops and no comparison is created; the app does not retry automatically.
- Comparison: local Python logic; no Gemini request.
- FX lookup: one non-Gemini request per processing batch when at least one vendor currency differs from the RFx comparison currency; no lookup when conversion is unnecessary.
- Analyst: one request for a new question and current comparison. Repeating the same question against unchanged data reuses the answer in the Streamlit session.

The app uses `gemini-3.5-flash-lite`, selected for low-cost multimodal document parsing and structured output.

The comparison currency is read from the RFx commercial terms. If none is provided, the buyer must enter a three-letter ISO currency code before processing. The comparison table displays comparable per-unit prices in that one currency. When vendor responses use other currencies, the app retrieves all required daily reference rates in one no-key Frankfurter request (not a Gemini call) and shows the applied rates and dates below the table. Original quotes remain available in each vendor's source document and extracted evidence. Totals are calculated only when currency, unit, and quantity basis are comparable; otherwise the price is flagged. These daily reference rates are indicative and may differ from a bank or supplier settlement rate.

The Analyst receives normalized extracted values and source evidence, not the original documents. Its guidance uses authoritative line counts and cautions against making an award recommendation without buyer weighting criteria; review important claims against the comparison and evidence.

If an Excel workbook's quote-table layout is not recognized, the app flags its line coverage for review instead of treating every RFx line as unquoted. Prices with an unstated or unsupported basis are marked not comparable with the reason shown in the comparison.

## Tests

Run the offline checks with:

```powershell
python -m unittest test_main.py
```

These tests do not use Gemini or consume API quota. A real extraction/Analyst run requires a valid Gemini key and available quota.
