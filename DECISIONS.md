# Build note: RFx Procurement Copilot

## What it does

The buyer enters instructions under User Reference: scope, line items, questionnaire, and terms. Gemini drafts an RFx from those instructions only; the buyer can review and edit it, then download it as a Word document. After at least one line item exists, the buyer can process up to five responses in one action. Quotes are aligned to the RFx and displayed side by side with questionnaire answers, commercial terms, source evidence, and exception flags. The Analyst can answer a new question from that current comparison and return a short answer with an optional table or chart.

## Key decisions

- Use local Python parsing for Excel and deterministic Python comparisons. This avoids sending spreadsheets to Gemini and avoids a second model call for routine arithmetic and mismatch checks.
- Batch unstructured files from one Process action into one Gemini request. Use the same low-cost multimodal model for RFx drafting, extraction, and Analyst answers; cache repeated Analyst questions for the same comparison in the current session.
- Preserve a short source excerpt and location with each extracted quote. Mark low-confidence mappings and unreadable source content for review rather than treating them as vendor omissions.
- Normalize clear unit aliases and explicit bases such as per 100 pieces. Use the buyer's RFx currency as the comparison currency and retrieve all required daily FX reference rates in one no-key Frankfurter request; show only normalized prices in the table and list applied rates/dates below it. Keep original quoted prices in the source documents and extracted evidence. Do not calculate converted prices when currency, rate, or unit is unknown or incompatible.
- Only normalize recognized price bases. For an unknown basis, show why the price is not comparable. If a spreadsheet quote layout is not recognized, flag the line coverage for review rather than presenting all requested lines as unquoted.
- The comparison requires the current buyer RFx so it can flag omissions and quantity mismatches against buyer requirements. Vendor responses do not need to use the buyer's template.
- Keep synthetic response files as local fixtures for extraction and comparison tests; the app UI does not expose a sample-demo action.

## Deliberately omitted

Supplier email sending, authentication, database storage, persistent audit logs, OCR services beyond Gemini vision, and production hosting are outside this take-home MVP. FX conversion uses indicative daily reference rates and adds no Gemini requests. The response comparison is session-local. A buyer should review source evidence and uncertain fields before making an award.

## Submission status

Live app: https://aerchain-procurement-copilot.streamlit.app/

Still needed: record the walkthrough and include the recording link with the live app link in the submission. The recording is not created by the app repository.
