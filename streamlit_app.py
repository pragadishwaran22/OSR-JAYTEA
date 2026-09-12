from __future__ import annotations

import hashlib

import pandas as pd
import streamlit as st
from streamlit.typing import UploadedFile

from branding import (
    APP_VERSION,
    COMPANY_NAME,
    LOGO_PATH,
    PREMIUM_BADGE,
    PRODUCT_NAME,
    VALUE_PROPOSITION,
    liquid_glass_css,
)
from report_engine import (
    ReportInputError,
    ReportValidationError,
    generate_workbook,
    prepare_report,
    preview_rows,
    safe_output_filename,
    validate_generated_workbook,
)


st.set_page_config(
    page_title=f"{PRODUCT_NAME} | {COMPANY_NAME}",
    page_icon=str(LOGO_PATH),
    layout="wide",
)

st.html(liquid_glass_css())

st.session_state.setdefault("generated_report", None)
st.session_state.setdefault("uploaded_fingerprint", None)

with st.container(key="brand_nav"):
    with st.container(
        horizontal=True,
        vertical_alignment="center",
        wrap=False,
        key="brand_nav_inner",
    ):
        st.image(LOGO_PATH, width=54)
        st.markdown(f"**{COMPANY_NAME}** Order intelligence")
        st.space("stretch")
        st.badge(f"Version {APP_VERSION}", color="yellow")

with st.container(key="hero"):
    hero_logo, hero_copy = st.columns([1.25, 4], gap="large", vertical_alignment="center")
    with hero_logo:
        st.image(LOGO_PATH, width=230)
    with hero_copy:
        st.badge(PREMIUM_BADGE, color="yellow", icon=":material/verified:")
        st.title(PRODUCT_NAME)
        st.markdown(VALUE_PROPOSITION)
        st.caption(
            ":material/lock: Your ERP workbook is processed locally in memory and is never modified."
        )

with st.container(key="workflow_intro"):
    st.subheader("Create your planning report", icon=":material/auto_awesome:")
    st.caption("Upload the daily workbook, choose the view, and generate a validated Excel file.")

upload_column, output_column = st.columns(2, gap="large")
with upload_column:
    with st.container(border=True, key="upload_panel"):
        st.subheader("1. Upload the daily report", icon=":material/upload_file:")
        st.caption("Select the original ERP Order Status workbook in XLSX format.")
        uploaded_file: UploadedFile | None = st.file_uploader(
            "Order Status workbook",
            type="xlsx",
            key="order_status_file",
            help="The workbook must contain the sheet 'Order Status-By shipment Date'.",
            max_upload_size=200,
        )
        st.caption(":material/database: Source values stay unchanged throughout processing.")

with output_column:
    with st.container(border=True, key="output_panel"):
        st.subheader("2. Choose the output", icon=":material/tune:")
        st.caption("Create the full workbook or focus on one planning condition.")
        mode_labels = {
            "Complete workbook": "all",
            "Highlighted report": "highlighted",
            "Red only": "red",
            "Black only": "black",
            "Yellow only": "yellow",
            "Green only": "green",
        }
        selected_label = st.segmented_control(
            "Report type",
            options=list(mode_labels),
            default="Complete workbook",
            required=True,
            key="report_type",
            wrap=True,
            width="stretch",
        )
        mode = mode_labels[selected_label]
        st.caption(
            "Complete workbook includes the highlighted report, four condition tabs, "
            "STD exclusion log, and run summary."
        )

if uploaded_file is not None:
    uploaded_bytes = uploaded_file.getvalue()
    fingerprint = hashlib.sha256(uploaded_bytes).hexdigest()
    if fingerprint != st.session_state.uploaded_fingerprint:
        st.session_state.uploaded_fingerprint = fingerprint
        st.session_state.generated_report = None
else:
    uploaded_bytes = None
    st.session_state.uploaded_fingerprint = None
    st.session_state.generated_report = None

current_result = st.session_state.generated_report
if current_result is not None and current_result["mode"] != mode:
    st.session_state.generated_report = None

with st.container(key="action_bar"):
    action_copy, action_button = st.columns([3.8, 1.25], vertical_alignment="center")
    with action_copy:
        if uploaded_bytes is None:
            st.markdown("**Ready when you are**")
            st.caption("Upload a valid workbook to unlock report generation.")
        else:
            st.markdown(f"**{uploaded_file.name}**")
            st.caption("Workbook loaded securely. Generate whenever you’re ready.")
    with action_button:
        generate_clicked = st.button(
            "Generate report",
            type="primary",
            icon=":material/auto_awesome:",
            disabled=uploaded_bytes is None,
            width="stretch",
        )

if generate_clicked and uploaded_bytes is not None:
    try:
        with st.status("Checking the workbook", expanded=True) as status:
            prepared = prepare_report(uploaded_bytes, uploaded_file.name)
            st.write(
                f"Found {prepared['order_count']} eligible orders and "
                f"{len(prepared['rows'])} non-STD FG lines."
            )
            status.update(label="Building the Excel report", state="running")
            workbook_bytes = generate_workbook(prepared, mode)
            status.update(label="Validating the result", state="running")
            validate_generated_workbook(workbook_bytes, prepared, mode)
            status.update(label="Report ready", state="complete", expanded=False)
        st.session_state.generated_report = {
            "mode": mode,
            "source_name": uploaded_file.name,
            "prepared": prepared,
            "workbook_bytes": workbook_bytes,
        }
        st.toast("The Excel report is ready.", icon=":material/check_circle:")
    except (ReportInputError, ReportValidationError) as exc:
        st.session_state.generated_report = None
        st.error(str(exc), icon=":material/error:")
    except Exception:
        st.session_state.generated_report = None
        st.error(
            "The report could not be generated. Confirm that the uploaded file is a valid ERP Order Status workbook.",
            icon=":material/error:",
        )

result = st.session_state.generated_report
if result is not None:
    prepared = result["prepared"]
    with st.container(key="result_shell"):
        st.badge("Validated output", color="green", icon=":material/verified:")
        st.subheader("Validation summary", icon=":material/fact_check:")
        metrics = st.container(horizontal=True, horizontal_alignment="distribute")
        metrics.metric("Eligible orders", prepared["order_count"], border=True)
        metrics.metric("Non-STD FG lines", len(prepared["rows"]), border=True)
        metrics.metric("Excluded STD lines", len(prepared["excluded_std"]), border=True)
        metrics.metric("All-STD orders removed", len(prepared["all_std_order_ids"]), border=True)

        condition_data = pd.DataFrame(
            [
                {"Condition": label, "FG lines": prepared["condition_counts"][label]}
                for label in ("Red", "Black", "Yellow", "Green", "No Fill")
            ]
        )
        preview_data = pd.DataFrame(preview_rows(prepared))
        summary_tab, preview_tab, exclusions_tab = st.tabs(
            [
                ":material/palette: Condition totals",
                ":material/preview: FG preview",
                ":material/block: STD exclusions",
            ]
        )
        with summary_tab:
            st.dataframe(
                condition_data,
                hide_index=True,
                width="content",
                column_config={
                    "FG lines": st.column_config.NumberColumn(format="%d"),
                },
            )
            st.caption("No Fill lines remain uncoloured in the Excel report.")
        with preview_tab:
            st.dataframe(
                preview_data,
                hide_index=True,
                height=430,
                key="fg_preview",
                column_config={
                    "Buyer Requested Shipment Date": st.column_config.DateColumn(format="DD-MMM-YYYY"),
                    "Pending Prod": st.column_config.NumberColumn(format="localized"),
                    "Planned Qty Total": st.column_config.NumberColumn(format="localized"),
                    "Un Planned": st.column_config.NumberColumn(format="localized"),
                },
            )
        with exclusions_tab:
            exclusions = pd.DataFrame(prepared["excluded_std"])
            if exclusions.empty:
                st.caption("No STD FG lines were found in the selected scope.")
            else:
                exclusions = exclusions.rename(
                    columns={
                        "output_order_seq": "Output Order Seq",
                        "source_order_seq": "Source Order Seq",
                        "order_id": "Order ID",
                        "source_row": "Source Row",
                        "fg_id": "FG Item ID",
                        "fg_name": "Prod Name",
                        "exclusion_type": "Exclusion Type",
                    }
                )
                st.dataframe(exclusions, hide_index=True, height=360, key="std_exclusions")

        st.download_button(
            "Download Excel report",
            data=result["workbook_bytes"],
            file_name=safe_output_filename(result["source_name"], result["mode"]),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            icon=":material/download:",
            on_click="ignore",
            width="content",
        )

with st.expander("Rules used by this app", icon=":material/rule:"):
    st.markdown(
        """
- Select the first **36 eligible orders** in source order.
- Remove FG lines whose `Prod Name` contains `STD`.
- Remove an entire order only when all its FG lines contain `STD`; replace it with the next eligible order.
- Treat the first two detected `Planned in <week>` columns as the acceptable planning period.
- Highlight only the planning section through `Un Planned` using red, black, yellow, or green.
- Do not perform BOM, stock, capacity, or planner-reason analysis.
"""
    )

st.caption(f"{COMPANY_NAME} • {PRODUCT_NAME} v{APP_VERSION} • Local processing")
