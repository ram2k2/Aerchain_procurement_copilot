from __future__ import annotations

import hashlib
import io
import json
from copy import deepcopy
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document

from main import analyst_answer, draft_rfx, empty_rfx, load_rfx, process_vendor_responses

st.set_page_config(page_title="RFx Procurement Copilot", page_icon="📦", layout="wide")
st.title("📦 RFx Procurement Copilot")
st.caption("Draft an RFx, compare vendor responses, and inspect the evidence behind an award decision.")


def _make_rfx_docx(rfx: dict) -> bytes:
    doc = Document()
    doc.add_heading("Request for Quotation", 0)
    doc.add_heading("Scope", level=1)
    doc.add_paragraph(str(rfx.get("scope", "")))
    doc.add_heading("Line items", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Shading Accent 1"
    for cell, value in zip(table.rows[0].cells, ["#", "Item", "Specification", "Quantity", "Unit"]):
        cell.text = value
    for idx, item in enumerate(rfx.get("items", []), 1):
        row = table.add_row().cells
        values = [item.get("item_id", idx), item.get("item", ""), item.get("specification", ""), item.get("quantity", ""), item.get("unit", "")]
        for cell, value in zip(row, values):
            cell.text = str(value or "")
    doc.add_heading("Supplier questionnaire", level=1)
    questions = rfx.get("questionnaire", [])
    if isinstance(questions, dict):
        questions = list(questions.values())
    for question in questions:
        doc.add_paragraph(str(question), style="List Number")
    doc.add_heading("Commercial terms", level=1)
    for key, value in (rfx.get("commercial_terms") or {}).items():
        doc.add_paragraph(f"{key.replace('_', ' ').title()}: {value}", style="List Bullet")
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def _cell(line: dict) -> str:
    if line["quote_status"] != "Quoted":
        return line["quote_status"]
    price = line.get("unit_price")
    currency = line.get("currency") or "?"
    unit = line.get("quoted_unit") or "unit unclear"
    basis = line.get("price_basis") or f"per {unit}"
    price_text = f"{currency} {price:g}" if isinstance(price, (int, float)) else "Price unclear"
    qty = line.get("quoted_quantity")
    if isinstance(qty, (int, float)) and isinstance(line.get("quantity"), (int, float)):
        qty_text = f" · Qty {qty:g}/{line['quantity']:g} {line['unit']}"
    elif isinstance(qty, (int, float)):
        qty_text = f" · Qty {qty:g} {line.get('quoted_unit') or ''} quoted"
    else:
        qty_text = ""
    converted = line.get("normalized_price")
    if currency != "INR" and converted is not None:
        price_text += f" ≈ INR {converted:g} (buyer rate)"
    confidence = line.get("confidence", "")
    return f"{price_text} ({basis}){qty_text}" + (f" · {confidence} confidence" if confidence else "")


def _side_by_side(results: list[dict], rfx: dict) -> pd.DataFrame:
    rows = []
    for i, req in enumerate(rfx["items"]):
        row = {"#": req["item_id"], "RFx item": req["item"], "Specification": req["specification"],
               "Required qty": req["quantity"], "Unit": req["unit"]}
        for result in results:
            row[result["vendor"]] = _cell(result["comparison"]["items"][i])
        rows.append(row)
    return pd.DataFrame(rows)


def _answer_matrix(results: list[dict], rfx: dict, field: str, label: str) -> pd.DataFrame:
    if field == "questionnaire":
        rows = [{label: question, **{result["vendor"]: result["terms"][field].get(question, "Not answered")
                                     or "Not answered"
                                   for result in results}}
                for question in rfx.get("questionnaire", [])]
    else:
        terms = rfx.get("commercial_terms") or {}
        rows = [{label: term.replace("_", " ").title(),
                 **{result["vendor"]: result["terms"][field].get(term, "Not stated") or "Not stated" for result in results}}
                for term in terms]
    return pd.DataFrame(rows)


def _results_key(results: list[dict], rfx: dict) -> str:
    payload = {"rfx": rfx, "vendors": [{"vendor": x["vendor"], "comparison": x["comparison"], "terms": x["terms"]} for x in results]}
    return hashlib.sha256(json.dumps(payload, sort_keys=True, default=str).encode()).hexdigest()


if st.session_state.get("_rfx_flow_version") != 2:
    # Clear any RFx or comparison from the earlier sample-prefilled flow.
    for key in ("rfx", "vendor_results", "processed_rfx", "vendor_sources", "analyst_cache",
                "analyst_response", "analyst_history", "gemini_extract_calls", "rfx_scope_edit", "rfx_items_edit",
                "rfx_questions_edit", "buyer_brief", "rfx_terms_edit", "rfx_drafted"):
        st.session_state.pop(key, None)
    st.session_state.rfx = empty_rfx()
    st.session_state.rfx_drafted = False
    st.session_state._rfx_flow_version = 2
rfx = st.session_state.rfx

with st.sidebar:
    st.header("Current RFx")
    st.write(rfx.get("scope") or "No RFx drafted in this session.")
    st.caption(f"{len(rfx.get('items', []))} RFx lines · {len(rfx.get('questionnaire', []))} supplier questions")
    st.caption("Spreadsheets and comparisons run locally. Gemini is used for drafting, unstructured extraction, and Analyst questions.")

tab1, tab2, tab3, tab4 = st.tabs(["Create RFx", "Vendor Responses", "Comparison", "AI Analyst"])

with tab1:
    st.header("Create RFx")
    st.subheader("User Reference")
    st.write("The user can provide the following instructions:")
    st.markdown("""
    - **Scope**
    - **Line items:** Item ID, Item, Specification, Quantity, Unit
    - **Questionnaire**
    - **Terms**
    """)
    buyer_input = st.text_area("Buyer instructions", placeholder="Enter the sourcing instructions using the sections above. Include only facts and requirements that should appear in the RFx.", key="buyer_brief", height=190)
    if st.button("✨ Generate RFx", type="primary"):
        if not buyer_input.strip():
            st.warning("Enter the buyer's RFx instructions first.")
        else:
            try:
                with st.spinner("Drafting the RFx (one Gemini call)…"):
                    st.session_state.rfx = draft_rfx(buyer_input.strip())
                st.session_state.rfx_drafted = True
                for key in ("vendor_results", "processed_rfx", "vendor_sources", "analyst_response", "analyst_history", "analyst_cache"):
                    st.session_state.pop(key, None)
                st.session_state.pop("rfx_scope_edit", None)
                st.session_state.pop("rfx_items_edit", None)
                st.session_state.pop("rfx_questions_edit", None)
                st.session_state.pop("rfx_terms_edit", None)
                st.success("RFx draft created. Review the scope, items, questions, and terms before sharing.")
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
    rfx = st.session_state.rfx
    if st.session_state.get("rfx_drafted"):
        st.divider()
        st.subheader("Generated RFx · Review and edit before sharing")
        rfx["scope"] = st.text_area("Scope", value=rfx.get("scope", ""), key="rfx_scope_edit")
        items_df = pd.DataFrame(rfx.get("items", []), columns=["item_id", "item", "specification", "quantity", "unit"] if not rfx.get("items") else None)
        rfx["items"] = st.data_editor(items_df, num_rows="dynamic", use_container_width=True, hide_index=True, key="rfx_items_edit").fillna("").to_dict("records")
        left, right = st.columns(2)
        with left:
            questions = rfx.get("questionnaire", [])
            if isinstance(questions, dict):
                questions = list(questions.values())
            question_text = st.text_area("Supplier questionnaire (one question per line)", value="\n".join(map(str, questions)), key="rfx_questions_edit", height=170)
            rfx["questionnaire"] = [x.strip() for x in question_text.splitlines() if x.strip()]
        with right:
            terms = pd.DataFrame([{"term": k.replace("_", " ").title(), "requirement": v}
                                  for k, v in (rfx.get("commercial_terms") or {}).items()],
                                 columns=["term", "requirement"])
            edited_terms = st.data_editor(terms, num_rows="dynamic", use_container_width=True, hide_index=True,
                                          key="rfx_terms_edit", column_config={"term": "Term", "requirement": "Requirement"})
            rfx["commercial_terms"] = {str(row["term"]).strip().lower().replace(" ", "_"): row["requirement"]
                                       for row in edited_terms.fillna("").to_dict("records") if str(row["term"]).strip()}
    st.download_button("📥 Download generated RFx (.docx)", data=_make_rfx_docx(rfx), file_name="RFx-draft.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       disabled=not st.session_state.get("rfx_drafted", False))

with tab2:
    st.header("Upload Vendor Responses")
    has_rfx_items = bool(st.session_state.rfx.get("items"))
    if not has_rfx_items:
        with st.container(border=True):
            st.warning("RFx document required")
            st.write("Upload the RFx document first to evaluate vendor responses against the requested requirements.")
            st.markdown("**① Upload RFx** → ② Upload Vendor Responses → ③ Compare & Evaluate")
            rfx_file = st.file_uploader("Upload RFx (.docx)", type=["docx"], key="rfx_upload")
            if rfx_file is not None:
                try:
                    uploaded_rfx = load_rfx(rfx_file.getvalue())
                    if not uploaded_rfx.get("items"):
                        raise ValueError("This RFx document has no readable line items.")
                    st.session_state.rfx = uploaded_rfx
                    st.session_state.rfx_drafted = True
                    for key in ("vendor_results", "processed_rfx", "vendor_sources", "analyst_response",
                                "analyst_history", "analyst_cache"):
                        st.session_state.pop(key, None)
                    for key in ("rfx_scope_edit", "rfx_items_edit", "rfx_questions_edit", "rfx_terms_edit"):
                        st.session_state.pop(key, None)
                    st.success(f"RFx uploaded with {len(uploaded_rfx['items'])} line items.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not read this RFx: {exc}")
    st.write("Upload up to five responses in their original formats. No vendor template is required.")
    uploaded = []
    for i in range(1, 6):
        file = st.file_uploader(f"Vendor {i}", type=["xlsx", "xls", "csv", "pdf", "docx", "jpg", "jpeg", "png", "webp", "txt", "eml"], key=f"vendor_{i}")
        if file is not None:
            uploaded.append(file)
    st.caption("One Gemini request extracts all non-spreadsheet files in this batch. Spreadsheets are read locally; comparisons use no Gemini call. Currency conversion is not applied.")
    if not has_rfx_items and uploaded:
        st.info(f"{len(uploaded)} vendor response(s) staged. Upload the RFx to enable comparison.")
    process_clicked = st.button("🔍 Process Vendor Responses", type="primary", disabled=not uploaded or not has_rfx_items)
    if process_clicked:
        st.session_state.vendor_results = []
        st.session_state.analyst_cache = {}
        st.session_state.analyst_history = []
        try:
            with st.spinner("Reading vendor responses and building the comparison…"):
                results, call_count = process_vendor_responses(uploaded, st.session_state.rfx)
            st.session_state.vendor_results = results
            st.session_state.processed_rfx = deepcopy(st.session_state.rfx)
            st.session_state.gemini_extract_calls = call_count
            st.session_state.analyst_cache = {}
            st.session_state.vendor_sources = {Path(getattr(source, "name", source)).name: source.getvalue() if hasattr(source, "getvalue") else Path(source).read_bytes() for source in uploaded}
            st.success(f"Processed {len(results)} response(s) using {call_count} Gemini extraction request(s).")
        except Exception as exc:
            st.error(f"Could not process this batch: {exc}")
    results = st.session_state.get("vendor_results", [])
    if results:
        st.markdown("**Processed vendors:**\n\n" + "\n".join(f"- {result['vendor']}" for result in results))

with tab3:
    st.header("Vendor Comparison")
    results = st.session_state.get("vendor_results", [])
    if not results:
        st.info("Upload and process vendor responses to see the side-by-side comparison.")
    else:
        comparison_rfx = st.session_state.get("processed_rfx", st.session_state.rfx)
        if comparison_rfx != st.session_state.rfx:
            st.info("The RFx was edited after this comparison was processed. Reprocess vendors to compare against the latest draft.")
        frame = _side_by_side(results, comparison_rfx)
        st.dataframe(frame, use_container_width=True, hide_index=True)
        st.download_button("Download comparison (CSV)", frame.to_csv(index=False).encode("utf-8-sig"), "vendor-comparison.csv", "text/csv")
        source_columns = st.columns(len(results))
        for column, result in zip(source_columns, results):
            source_bytes = st.session_state.get("vendor_sources", {}).get(result["file_name"])
            if source_bytes:
                column.download_button(f"📎 {result['vendor']} document", source_bytes,
                                       file_name=result["file_name"], key=f"source_{result['file_name']}")
            else:
                column.caption(f"{result['vendor']}: original file unavailable")
        questionnaire_frame = _answer_matrix(results, comparison_rfx, "questionnaire", "Question")
        if not questionnaire_frame.empty:
            st.subheader("Questionnaire answers")
            st.dataframe(questionnaire_frame, use_container_width=True, hide_index=True)
        terms_frame = _answer_matrix(results, comparison_rfx, "commercial_terms", "Term")
        if not terms_frame.empty:
            st.subheader("Commercial terms")
            st.dataframe(terms_frame, use_container_width=True, hide_index=True)
        for result in results:
            comp = result["comparison"]
            with st.expander(f"{result['vendor']} · flagged items and source evidence"):
                st.write("**Lines not quoted in a readable response**", comp["missing_items"])
                st.write("**Quantity mismatches**", comp["quantity_mismatches"])
                st.write("**Unit mismatches**", comp["unit_mismatches"])
                st.write("**Currency mismatches**", comp["currency_mismatches"])
                st.write("**Other issues / evidence**", comp["other_issues"])
                st.write("**Specification differences**", comp["specification_mismatches"])
                st.write("**Questionnaire gaps**", comp["questionnaire_issues"])
                st.write("**Commercial-term review**", comp["term_issues"])
                evidence_rows = [{"RFx item": line["item"], "status": line["quote_status"],
                                  "evidence": line.get("evidence", ""), "source location": line.get("source_location", ""),
                                  "extraction notes": line.get("notes", "")}
                                 for line in comp["items"] if line.get("evidence") or line.get("notes")]
                if evidence_rows:
                    st.write("**Extracted evidence by line**")
                    st.dataframe(pd.DataFrame(evidence_rows), use_container_width=True, hide_index=True)

with tab4:
    st.header("💬 AI Analyst")
    results = st.session_state.get("vendor_results", [])
    history = st.session_state.setdefault("analyst_history", [])

    def _render_analyst_answer(answer: dict, answer_id: int) -> None:
        st.markdown(str(answer.get("answer", "")))
        table = answer.get("table")
        if isinstance(table, dict) and isinstance(table.get("rows"), list) and table.get("rows"):
            st.markdown(f"**{table.get('title') or 'Supporting data'}**")
            try:
                table_df = pd.DataFrame(table["rows"], columns=table.get("columns") or None)
                st.dataframe(table_df, use_container_width=True, hide_index=True)
                st.download_button("Download answer table (CSV)", table_df.to_csv(index=False).encode("utf-8-sig"),
                                   "analyst-table.csv", "text/csv", key=f"analyst_table_{answer_id}")
            except (ValueError, TypeError) as exc:
                st.warning(f"The answer table could not be displayed: {exc}")
        chart = answer.get("chart")
        if isinstance(chart, dict) and chart.get("labels") and len(chart.get("labels", [])) == len(chart.get("values", [])):
            chart_values = pd.to_numeric(pd.Series(chart["values"]), errors="coerce")
            if chart_values.notna().all():
                st.markdown(f"**{chart.get('title') or 'Supporting chart'}**")
                st.bar_chart(pd.Series(chart_values.tolist(), index=chart["labels"]))
            else:
                st.warning("The Analyst chart contained non-numeric values and was omitted.")
        st.caption("Verify source evidence and flagged uncertainty before making an award decision.")

    for message_index, message in enumerate(history):
        with st.chat_message(message["role"]):
            if message["role"] == "user":
                st.markdown(message["content"])
            elif message.get("response"):
                _render_analyst_answer(message["response"], message_index)
            else:
                st.markdown(message.get("content", ""))

    question = st.chat_input("Ask a question about the vendor responses…", disabled=not results)
    if question and question.strip():
        question = question.strip()
        history.append({"role": "user", "content": question})
        comparison_rfx = st.session_state.get("processed_rfx", st.session_state.rfx)
        dataset_key = _results_key(results, comparison_rfx)
        cache = st.session_state.setdefault("analyst_cache", {})
        cache_key = f"{dataset_key}:{question.casefold()}"
        try:
            if cache_key in cache:
                response = cache[cache_key]
            else:
                with st.spinner("Analyzing the current comparison (one Gemini call)…"):
                    response = analyst_answer(question, results, comparison_rfx)
                cache[cache_key] = response
            history.append({"role": "assistant", "response": response})
            st.rerun()
        except Exception as exc:
            st.error(str(exc))
    if not results:
        st.caption("Process vendor responses first to start a conversation about the comparison.")
    st.divider()
    st.subheader("Example questions")
    st.write("• Which vendor is cheapest on comparable lines?  • Which items are missing?  • What if we split the award by line?  • Which quality answers are unknown?")
