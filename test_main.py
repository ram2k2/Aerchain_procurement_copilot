import unittest
import sys
import io
from unittest.mock import MagicMock, patch
from types import ModuleType, SimpleNamespace
from pathlib import Path

import main
import pandas as pd

ROOT = Path(__file__).resolve().parent


def fake_genai_modules():
    google_module = ModuleType("google")
    google_module.__path__ = []
    genai_module = ModuleType("google.genai")
    genai_module.types = SimpleNamespace(GenerateContentConfig=lambda **kwargs: kwargs)
    google_module.genai = genai_module
    return {"google": google_module, "google.genai": genai_module}


class RfxTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rfx = main.load_rfx()

    def test_reference_has_expected_schedule_and_questions(self):
        self.assertEqual(len(self.rfx["items"]), 30)
        self.assertEqual(len(self.rfx["questionnaire"]), 7)
        self.assertIn("45 days", self.rfx["commercial_terms"]["payment_terms"])

    def test_comparison_currency_is_taken_from_rfx_without_an_inr_default(self):
        self.assertEqual(main.infer_rfx_currency(self.rfx), "INR")
        self.assertEqual(main.infer_rfx_currency({"commercial_terms": {}}), "")
        self.assertEqual(main.infer_rfx_currency({"commercial_terms": {"currency": "Prices quoted in EUR"}}), "EUR")
        no_currency = dict(self.rfx, commercial_terms={})
        with self.assertRaisesRegex(ValueError, "Set a comparison currency"):
            main.process_vendor_responses([ROOT / "Vendor_1.xlsx"], no_currency)

    def test_latest_fx_rates_are_fetched_together_and_inverted_for_conversion(self):
        class Response:
            def __enter__(self):
                return self
            def __exit__(self, *_args):
                return False
            def read(self):
                return (b'[{"date":"2026-09-25","base":"INR","quote":"USD","rate":0.012},'
                        b'{"date":"2026-09-25","base":"INR","quote":"EUR","rate":0.011}]')

        with patch("urllib.request.urlopen", return_value=Response()) as request:
            rates, dates, error = main._fetch_fx_rates("INR", ["USD", "EUR", "INR"])
        request.assert_called_once()
        self.assertIn("base=INR", request.call_args.args[0].full_url)
        self.assertIn("quotes=EUR%2CUSD", request.call_args.args[0].full_url)
        self.assertAlmostEqual(rates["USD"], 1 / 0.012)
        self.assertAlmostEqual(rates["EUR"], 1 / 0.011)
        self.assertEqual(dates, {"USD": "2026-09-25", "EUR": "2026-09-25"})
        self.assertIsNone(error)

    def test_conversion_uses_selected_comparison_currency_and_keeps_original_quote(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "pcs",
            "unit_price": 100, "currency": "USD", "price_basis": "per piece", "confidence": "high",
            "evidence": "USD 100 per piece",
        }]}
        comparison = main.compare_vendor_with_rfx(
            vendor, self.rfx, {"USD": 0.92}, base_currency="EUR", fx_rate_dates={"USD": "2026-09-25"}
        )
        row = comparison["items"][0]
        self.assertEqual(row["currency"], "USD")
        self.assertEqual(row["unit_price"], 100)
        self.assertAlmostEqual(row["normalized_price"], 92)
        self.assertAlmostEqual(row["comparable_total"], 460000)
        self.assertEqual(row["comparison_currency"], "EUR")
        self.assertEqual(row["fx_rate_date"], "2026-09-25")

    def test_unavailable_fx_rate_keeps_original_price_and_flags_it(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "pcs",
            "unit_price": 100, "currency": "USD", "price_basis": "per piece", "confidence": "high",
        }]}
        row = main.compare_vendor_with_rfx(vendor, self.rfx, {}, base_currency="EUR")["items"][0]
        self.assertEqual(row["unit_price"], 100)
        self.assertEqual(row["currency"], "USD")
        self.assertIsNone(row["normalized_price"])
        self.assertIsNone(row["comparable_total"])

    def test_currency_conversion_does_not_claim_a_price_in_an_incompatible_unit(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "kg",
            "unit_price": 100, "currency": "USD", "price_basis": "per kg", "confidence": "high",
        }]}
        comparison = main.compare_vendor_with_rfx(vendor, self.rfx, {"USD": 83}, base_currency="INR")
        row = comparison["items"][0]
        self.assertEqual(row["currency"], "USD")
        self.assertIsNone(row["normalized_price"])
        self.assertIsNone(row["comparable_total"])
        self.assertEqual(comparison["unit_mismatches"][0]["quoted"], "kg")

    def test_gemini_client_reads_streamlit_cloud_secret(self):
        modules = fake_genai_modules()
        client_constructor = MagicMock(return_value=object())
        modules["google.genai"].Client = client_constructor
        streamlit_module = ModuleType("streamlit")
        streamlit_module.secrets = {"GEMINI_API_KEY": "cloud-test-key"}
        with patch.dict(sys.modules, {**modules, "streamlit": streamlit_module}), patch.dict(
            "os.environ", {"GEMINI_API_KEY": "", "GOOGLE_API_KEY": ""}
        ):
            main._client()
        client_constructor.assert_called_once_with(api_key="cloud-test-key")

    def test_uploaded_rfx_docx_bytes_are_parsed(self):
        uploaded = main.load_rfx((ROOT / "RFx.docx").read_bytes())
        self.assertEqual(len(uploaded["items"]), 30)
        self.assertEqual(len(uploaded["questionnaire"]), 7)
        self.assertIn("45 days", uploaded["commercial_terms"]["payment_terms"])

    def test_new_rfx_starts_empty_and_vendor_processing_requires_one(self):
        self.assertEqual(main.empty_rfx(), {"scope": "", "items": [], "questionnaire": [], "commercial_terms": {}})
        with self.assertRaisesRegex(ValueError, "Generate an RFx"):
            main.process_vendor_responses([ROOT / "Vendor_1.xlsx"])

    def test_vendor_questionnaire_is_matched_to_buyer_defined_questions(self):
        custom_rfx = {"scope": "Test", "items": [], "questionnaire": ["Can you meet the custom SLA?"],
                      "commercial_terms": {}}
        vendor = main._normalise({"questionnaire": {"Can you meet the custom SLA?": "Yes"}}, "vendor.txt", custom_rfx)
        self.assertEqual(vendor["questionnaire"], {"Can you meet the custom SLA?": "Yes"})

    def test_missing_vendor_currency_stays_unknown(self):
        vendor = {"source_status": "readable", "items": [{
            "item": self.rfx["items"][0]["item"], "rfx_line_id": 1, "quantity": 1,
            "unit": self.rfx["items"][0]["unit"], "unit_price": 5, "confidence": "high",
        }]}
        row = main.compare_vendor_with_rfx(vendor, self.rfx)["items"][0]
        self.assertEqual(row["currency"], "Unknown")
        self.assertIsNone(row["comparable_total"])

    def test_comparison_accepts_text_quantity_in_edited_rfx(self):
        edited_rfx = {"scope": "Test", "items": [{"item_id": "1", "item": "Boxes",
                      "specification": "Small box", "quantity": "5000", "unit": "pcs"}],
                      "questionnaire": [], "commercial_terms": {"currency": "INR"}}
        vendor = {"source_status": "readable", "items": [{
            "item": "Boxes", "rfx_line_id": 1, "quantity": 5000.0, "unit": "pcs",
            "unit_price": 2, "currency": "INR", "price_basis": "per pcs", "confidence": "high",
        }]}
        row = main.compare_vendor_with_rfx(vendor, edited_rfx)["items"][0]
        self.assertEqual(row["quantity"], 5000.0)
        self.assertEqual(row["quoted_quantity"], 5000.0)
        self.assertEqual(row["comparable_total"], 10000.0)

    def test_supplier_extraction_prompt_uses_buyer_questions(self):
        client = SimpleNamespace(models=SimpleNamespace(generate_content=MagicMock(
            return_value=SimpleNamespace(text='{"vendors":[]}'))))
        custom_rfx = {"scope": "Custom project", "items": [{"item_id": 1, "item": "Custom item"}],
                      "questionnaire": ["Can you meet our custom SLA?"], "commercial_terms": {}}
        with patch.dict(sys.modules, fake_genai_modules()), patch("main._client", return_value=client):
            main._extract_batch([], custom_rfx)
        prompt = client.models.generate_content.call_args.kwargs["contents"][0]
        self.assertIn("Can you meet our custom SLA?", prompt)
        self.assertIn("exactly 14 strings", prompt)
        self.assertIn("an empty string if uncertain", prompt)
        self.assertNotIn("iso_9001", prompt)
        config = client.models.generate_content.call_args.kwargs["config"]
        self.assertEqual(config["response_mime_type"], "application/json")
        self.assertEqual(config["max_output_tokens"], main.MAX_OUTPUT_TOKENS)
        self.assertEqual(main.MAX_OUTPUT_TOKENS, 65536)
        self.assertEqual(config["response_schema"]["properties"]["vendors"]["type"], "array")
        self.assertNotIn("additionalProperties", str(config["response_schema"]))
        item_schema = config["response_schema"]["properties"]["vendors"]["items"]["properties"]["items"]
        self.assertEqual(item_schema["items"]["type"], "array")
        vendor_schema = config["response_schema"]["properties"]["vendors"]["items"]["properties"]
        self.assertEqual(vendor_schema["questionnaire"]["type"], "array")
        self.assertIn("question", vendor_schema["questionnaire"]["items"]["properties"])
        self.assertEqual(vendor_schema["commercial_terms"]["type"], "array")

    def test_question_answer_and_term_response_arrays_normalize_to_rfx_keys(self):
        first_question, second_question = self.rfx["questionnaire"][:2]
        vendor = main._normalise({
            "vendor": "Pair Format Vendor", "source_status": "readable", "items": [],
            "questionnaire": [
                {"question": first_question.upper(), "answer": "Yes, certified."},
                {"question": second_question, "answer": ""},
            ],
            "commercial_terms": [
                {"term": "Quote Validity", "response": "60 days"},
                {"term": "Payment Terms", "response": "45 days from invoice"},
            ],
        }, "pair_format.txt", self.rfx)
        self.assertEqual(vendor["questionnaire"][first_question], "Yes, certified.")
        self.assertEqual(vendor["questionnaire"][second_question], "")
        self.assertEqual(vendor["commercial_terms"]["quote_validity"], "60 days")
        self.assertEqual(vendor["commercial_terms"]["payment_terms"], "45 days from invoice")

    def test_compact_vendor_item_rows_normalize_to_comparison_fields(self):
        vendor = main._normalise({
            "vendor": "Compact Vendor", "source_status": "readable", "items": [[
                "1", "3-Ply Small Box", "300x200x150 mm", "5000", "pcs", "17.28", "INR",
                "per pcs", "high", "INR 17.28 per piece", "page 1", "18%", "12", "",
            ]],
        }, "compact.txt", self.rfx)
        item = vendor["items"][0]
        self.assertEqual(item["rfx_line_id"], 1.0)
        self.assertEqual(item["quantity"], 5000.0)
        self.assertEqual(item["unit_price"], 17.28)
        self.assertEqual(item["evidence"], "INR 17.28 per piece")

    def test_malformed_vendor_json_fails_cleanly_without_an_automatic_retry(self):
        client = SimpleNamespace(models=SimpleNamespace(generate_content=MagicMock(
            return_value=SimpleNamespace(text='{"vendors":[{', candidates=[
                SimpleNamespace(finish_reason="MAX_TOKENS")
            ]))))
        with patch.dict(sys.modules, fake_genai_modules()), patch("main._client", return_value=client):
            with self.assertRaisesRegex(ValueError, "finish reason: MAX_TOKENS") as error:
                main._extract_batch([{"name": "Vendor.txt", "text": "sample quote"}], self.rfx)
        self.assertIn("Gemini JSON is incomplete", str(error.exception))
        self.assertIn("This request was not retried automatically", str(error.exception))
        client.models.generate_content.assert_called_once()

    def test_spreadsheet_parses_locally_and_aligns_all_rfx_rows(self):
        result, calls = main.process_vendor_responses([ROOT / "Vendor_1.xlsx"], self.rfx)
        self.assertEqual(calls, 0)
        self.assertEqual(len(result), 1)
        self.assertEqual(len(result[0]["comparison"]["items"]), 30)
        self.assertEqual(result[0]["comparison"]["items"][0]["quote_status"], "Quoted")
        self.assertEqual(len(result[0]["comparison"]["missing_items"]), 5)
        self.assertEqual(len(result[0]["comparison"]["quantity_mismatches"]), 1)
        self.assertEqual(result[0]["terms"]["commercial_terms"]["quote_validity"], "60 days")
        self.assertEqual(result[0]["terms"]["commercial_terms"]["payment_terms"], "45 days from invoice")
        self.assertEqual(result[0]["comparison"]["term_issues"], [])

    def test_supplied_pdf_and_docx_text_readers_include_vendor_content(self):
        pdf = main._prepare_unstructured("Vendor_2.pdf", (ROOT / "Vendor_2.pdf").read_bytes())
        docx = main._prepare_unstructured("Vendor_3.docx", (ROOT / "Vendor_3.docx").read_bytes())
        self.assertIn("ABC Packaging", pdf["text"])
        self.assertIn("XYZ Packaging", docx["text"])

    def test_photo_is_passed_as_image_part_in_the_single_batch(self):
        class Upload:
            name = "quote.png"
            def getvalue(self):
                return b"sample image bytes"

        def capture(documents, rfx):
            self.assertEqual(documents[0]["mime"], "image/png")
            self.assertEqual(documents[0]["bytes"], b"sample image bytes")
            return [{"file_name": "quote.png", "vendor": "Photo Vendor", "source_status": "unreadable", "items": []}]

        with patch("main._extract_batch", side_effect=capture) as extract:
            results, calls = main.process_vendor_responses([Upload()], self.rfx)
        extract.assert_called_once()
        self.assertEqual(calls, 1)
        self.assertEqual(results[0]["comparison"]["items"][0]["quote_status"], "Unable to read source")

    def test_unreadable_source_is_not_misreported_as_vendor_omission(self):
        vendor = {"source_status": "unreadable", "items": [], "source_notes": "Photo text could not be read."}
        comparison = main.compare_vendor_with_rfx(vendor, self.rfx)
        self.assertEqual(comparison["missing_items"], [])
        self.assertEqual({x["quote_status"] for x in comparison["items"]}, {"Unable to read source"})

    def test_unit_basis_and_buyer_fx_rate_normalize_price(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "pcs",
            "unit_price": 50, "currency": "USD", "price_basis": "per 100 pieces",
            "confidence": "high", "evidence": "USD 50 per 100 pieces",
        }]}
        row = main.compare_vendor_with_rfx(vendor, self.rfx, {"USD": 83})["items"][0]
        self.assertAlmostEqual(row["normalized_price"], 41.5)
        self.assertAlmostEqual(row["comparable_total"], 207500)

    def test_unsupported_price_basis_is_not_normalized_and_has_reason(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "pcs",
            "unit_price": 120, "currency": "INR", "price_basis": "per dozen pieces",
            "confidence": "high", "evidence": "INR 120 per dozen pieces",
        }]}
        row = main.compare_vendor_with_rfx(vendor, self.rfx)["items"][0]
        self.assertIsNone(row["normalized_price"])
        self.assertIn("per dozen", row["price_not_comparable_reason"])
        self.assertIn("cannot be converted", row["price_not_comparable_reason"])

    def test_unrecognized_spreadsheet_layout_is_review_not_vendor_omission(self):
        buffer = io.BytesIO()
        pd.DataFrame({"Product description": ["Small box"], "Amount": [15]}).to_excel(buffer, index=False)
        parsed = main._spreadsheet("unfamiliar.xlsx", buffer.getvalue())
        self.assertEqual(parsed["source_status"], "unrecognized_layout")
        vendor = main._normalise(parsed, "unfamiliar.xlsx", self.rfx)
        comparison = main.compare_vendor_with_rfx(vendor, self.rfx)
        self.assertEqual(comparison["missing_items"], [])
        self.assertEqual(comparison["items"][0]["quote_status"], "Needs review - quote table format not recognized")
        self.assertIn("Quote table format not recognized", comparison["vendor_notes"])

    def test_duplicate_rfx_names_use_specification_to_match_correct_line(self):
        vendor = {"source_status": "readable", "items": [{
            "item": "Corrugated Sheet", "specification": "1000x600 mm, 3-Ply", "quantity": 4000,
            "unit": "sheets", "unit_price": 72, "currency": "INR", "price_basis": "per sheet", "confidence": "high",
        }]}
        lines = main.compare_vendor_with_rfx(vendor, self.rfx)["items"]
        self.assertEqual(lines[16]["quote_status"], "Not quoted in readable response")
        self.assertEqual(lines[17]["quote_status"], "Quoted")

    def test_vendor_batch_uses_one_extraction_call_for_all_unstructured_files(self):
        synthetic = {
            "file_name": "Vendor_4.txt", "vendor": "Delta Cartons", "source_status": "readable",
            "items": [{"item": "3-Ply Small Box", "rfx_line_id": 1, "quantity": 5000, "unit": "pcs",
                       "unit_price": 0.22, "currency": "USD", "price_basis": "per piece", "confidence": "high",
                       "evidence": "USD 0.22 per piece"}],
            "questionnaire": {"iso_9001": "Yes"}, "commercial_terms": {"freight": "extra"},
        }
        with patch("main._extract_batch", return_value=[synthetic]) as extract:
            results, calls = main.process_vendor_responses(
                [ROOT / "Vendor_1.xlsx", ROOT / "Vendor_4.txt"], self.rfx, {"USD": 83}
            )
        extract.assert_called_once()
        self.assertEqual(calls, 1)
        self.assertEqual(len(results), 2)
        self.assertEqual(results[1]["vendor"], "Delta Cartons")
        self.assertEqual(results[1]["comparison"]["items"][0]["quote_status"], "Quoted")

    def test_five_vendor_demo_builds_a_30_line_comparison(self):
        def quote(line_id, *, quantity=None, unit_price=10, currency="INR", basis="per unit", item=None):
            req = self.rfx["items"][line_id - 1]
            return {"rfx_line_id": line_id, "item": item or req["item"], "specification": req["specification"],
                    "quantity": req["quantity"] if quantity is None else quantity, "unit": req["unit"],
                    "unit_price": unit_price, "currency": currency, "price_basis": basis,
                    "confidence": "high", "evidence": f"Quoted line {line_id}"}

        outputs = [
            {"file_name": "Vendor_2.pdf", "vendor": "ABC Packaging", "source_status": "readable",
             "items": [quote(1, unit_price=18), quote(2, unit_price=27)], "questionnaire": {"iso_9001": "Yes"},
             "commercial_terms": {"payment_terms": "45 days from invoice"}},
            {"file_name": "Vendor_3.docx", "vendor": "XYZ Packaging", "source_status": "readable",
             "items": [quote(4, unit_price=22), quote(20, unit_price=8)], "questionnaire": {}, "commercial_terms": {}},
            {"file_name": "Vendor_4.txt", "vendor": "Delta Cartons", "source_status": "readable",
             "items": [quote(i, quantity=3500 if i == 5 else None, unit_price=0.30 + i / 100,
                              currency="USD", basis="per piece") for i in [3, *range(5, 16)]],
             "questionnaire": {}, "commercial_terms": {}},
            {"file_name": "Vendor_5.png", "vendor": "Meridian Packaging", "source_status": "readable",
             "items": [quote(i, unit_price=1800 if i == 16 else 120 if i == 21 else 1400 if i == 24 else 2100 if i == 25 else 20,
                              basis="per 100 pieces" if i in {16, 21, 24, 25} else "per unit")
                       for i in [16, 18, 19, *range(21, 31)]],
             "questionnaire": {"defect_tolerance": "Not specified"},
             "commercial_terms": {"payment_terms": "45 days", "freight": "extra"},
             "source_notes": "Remaining item prices refer to an unattached prior-year quote."},
        ]
        paths = [ROOT / name for name in ("Vendor_1.xlsx", "Vendor_2.pdf", "Vendor_3.docx", "Vendor_4.txt", "Vendor_5.png")]
        outputs_by_name = {output["file_name"]: output for output in outputs}

        def extract_all(documents, _rfx):
            return [outputs_by_name[doc["name"]] for doc in documents]

        with patch("main._extract_batch", side_effect=extract_all) as extract, \
             patch("main._fetch_fx_rates", return_value=({"USD": 83.0}, {"USD": "2026-09-25"}, None)) as fetch_rates:
            results, calls = main.process_vendor_responses(paths, self.rfx)
        extract.assert_called_once()
        fetch_rates.assert_called_once_with("INR", ["USD"])
        self.assertEqual(calls, 1)
        self.assertEqual([doc["name"] for doc in extract.call_args.args[0]], [
            "Vendor_2.pdf", "Vendor_3.docx", "Vendor_4.txt", "Vendor_5.png"
        ])
        self.assertEqual(len(results), 5)
        self.assertTrue(all(len(result["comparison"]["items"]) == 30 for result in results))
        covered = {line["item_id"] for result in results for line in result["comparison"]["items"] if line["quote_status"] == "Quoted"}
        self.assertEqual(covered, set(range(1, 31)))
        delta = next(r for r in results if r["vendor"] == "Delta Cartons")
        self.assertTrue(delta["comparison"]["quantity_mismatches"])
        self.assertEqual(delta["comparison"]["currency_mismatches"], [])
        self.assertAlmostEqual(delta["comparison"]["items"][2]["normalized_price"], 0.33 * 83)
        meridian = next(r for r in results if r["vendor"] == "Meridian Packaging")
        self.assertAlmostEqual(meridian["comparison"]["items"][15]["normalized_price"], 18)

    def test_incomplete_batch_fails_closed_after_one_request(self):
        sources = [ROOT / name for name in ("Vendor_2.pdf", "Vendor_3.docx", "Vendor_4.txt")]
        incomplete_batch = [
            {"file_name": "Vendor_2.pdf", "vendor": "Vendor 2", "source_status": "readable",
             "items": [], "questionnaire": {}, "commercial_terms": {}},
            {"file_name": "Vendor_3.docx", "vendor": "Vendor 3", "source_status": "readable",
             "items": [], "questionnaire": {}, "commercial_terms": {}},
        ]
        with patch("main._extract_batch", return_value=incomplete_batch) as extract:
            with self.assertRaisesRegex(ValueError, "after 1 Gemini request") as error:
                main.process_vendor_responses(sources, self.rfx)
        extract.assert_called_once()
        self.assertIn("No comparison was created", str(error.exception))
        self.assertIn("vendor_4.txt", str(error.exception))

    def test_rfx_draft_is_one_real_model_request_and_json_is_parsed(self):
        client = SimpleNamespace(models=SimpleNamespace(generate_content=MagicMock(
            return_value=SimpleNamespace(text='{"scope":"Custom sourcing need","items":[{"item_id":"A-7","item":"Custom item","specification":"Buyer specification","quantity":12,"unit":"boxes"}],"questionnaire":["Can you meet the custom SLA?"],"commercial_terms":{"payment_terms":"Net 30"}}'))))
        with patch.dict(sys.modules, fake_genai_modules()), patch("main._client", return_value=client):
            draft = main.draft_rfx("Source custom items for the new site.")
        client.models.generate_content.assert_called_once()
        self.assertEqual(draft["scope"], "Custom sourcing need")
        self.assertEqual(draft["items"][0]["item_id"], "A-7")
        self.assertEqual(draft["questionnaire"], ["Can you meet the custom SLA?"])
        self.assertEqual(draft["commercial_terms"], {"payment_terms": "Net 30"})
        prompt = client.models.generate_content.call_args.kwargs["contents"]
        self.assertIn("Source custom items for the new site.", prompt)
        self.assertNotIn("Corrugated Sheet", prompt)

    def test_analyst_is_one_model_request_and_returns_renderable_views(self):
        results, _ = main.process_vendor_responses([ROOT / "Vendor_1.xlsx"], self.rfx)
        answer_json = '{"answer":"Vendor 1 quoted five lines.","table":{"title":"Coverage","columns":["Vendor","Lines"],"rows":[["Vendor_1",5]]},"chart":{"title":"Quoted lines","labels":["Vendor_1"],"values":[5]}}'
        client = SimpleNamespace(models=SimpleNamespace(generate_content=MagicMock(
            return_value=SimpleNamespace(text=answer_json))))
        with patch.dict(sys.modules, fake_genai_modules()), patch("main._client", return_value=client):
            answer = main.analyst_answer("How many lines did this vendor quote?", results, self.rfx)
        client.models.generate_content.assert_called_once()
        self.assertEqual(answer["table"]["rows"][0][1], 5)
        self.assertEqual(answer["chart"]["values"], [5])
        prompt = client.models.generate_content.call_args.kwargs["contents"]
        self.assertIn('"quote_status_counts": {"Quoted": 23, "Not quoted in readable response": 5, "Needs review": 2}', prompt)
        self.assertIn("only empty or null values are missing", prompt)
        self.assertIn("do not add generic caveats or unrelated categories", prompt)
        self.assertIn("Never claim a vendor is the only one to quote all lines", prompt)
        self.assertIn("do not imply a definitive award where no weighting criteria were supplied", prompt)


if __name__ == "__main__":
    unittest.main()
