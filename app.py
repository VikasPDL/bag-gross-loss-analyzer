import pandas as pd
import streamlit as st

from processing import (
    InvalidWorkbookError,
    build_category_mismatch_summary,
    build_category_weight_summary,
    build_department_summary,
    build_gross_loss_report,
    build_returned_bag_category_counts,
    build_returned_bag_details,
    to_excel_bytes,
)

build_gross_loss_report = st.cache_data(show_spinner=False)(build_gross_loss_report)
build_department_summary = st.cache_data(show_spinner=False)(build_department_summary)
build_returned_bag_details = st.cache_data(show_spinner=False)(build_returned_bag_details)

st.set_page_config(
    page_title="Bag Gross Loss % Analyzer",
    page_icon="\U0001F48D",
    layout="wide",
)

CUSTOM_CSS = """
<style>
#MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 2rem; padding-bottom: 3rem;}
[data-testid="stMetric"] {
    background: rgba(48, 84, 150, 0.08);
    border: 1px solid rgba(48, 84, 150, 0.18);
    border-radius: 10px;
    padding: 14px 16px 8px 16px;
}
[data-testid="stMetricLabel"] {font-weight: 600;}
h1 {font-weight: 700;}
.stTabs [data-baseweb="tab-list"] {gap: 4px;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.title("\U0001F48D Bag Gross Loss % Analyzer")
st.caption(
    "Upload the Bag Transaction export to compute Opening / Add Metal / Return Metal / "
    "Loss Metal / Balance and Gross Loss % for every bag, one row per bag."
)

with st.sidebar:
    st.header("Upload")
    uploaded_files = st.file_uploader(
        "Bag Transaction Excel file(s) (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more series exports (e.g. 25, 26) — they'll be combined into one report.",
    )
    st.divider()
    st.header("Settings")
    metal_raw_name = st.selectbox(
        "Raw Name value treated as Metal",
        options=["Gold", "TYPE2"],
        index=0,
        help="Only rows tagged with this Raw Name are used for Add Metal / Return Metal / Loss Metal.",
    )
    mismatch_tolerance = st.number_input(
        "Balance vs Last Weight tolerance (gm)",
        min_value=0.0,
        value=0.001,
        step=0.001,
        format="%.3f",
    )
    exclude_recast = st.checkbox(
        "Exclude returned/recast bags",
        value=True,
        help="Leaves out bags that went through CAST more than once (melted back and recast) from "
        "the Report, Category Summary, Department Summary and Charts. They're still listed, by "
        "category, in the Returned/Recast Bags tab.",
    )
    st.divider()
    st.caption("Bag No is rebuilt as Year/Bag Type/Bag No, e.g. 26/G/8518.")

if not uploaded_files:
    st.info("⬆️ Upload one or more Bag Transaction .xlsx files from the sidebar to get started.")
    st.stop()

total_size_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
if total_size_mb > 30:
    st.caption(f"⏳ {total_size_mb:.0f} MB uploaded — large files can take a few minutes to process.")

try:
    with st.spinner(f"Processing {len(uploaded_files)} workbook(s)..."):
        df = build_gross_loss_report(
            uploaded_files,
            metal_raw_name=metal_raw_name,
            mismatch_tolerance=mismatch_tolerance,
            exclude_recast=exclude_recast,
        )
except InvalidWorkbookError as exc:
    st.error(f"⚠️ {exc}")
    st.stop()
except Exception as exc:
    st.error(
        "⚠️ Something went wrong while reading one of the uploaded files. This app expects the "
        "'Bag Transaction' detail export — if this file is a different report, or is corrupted, "
        "that's likely why."
    )
    with st.expander("Technical details"):
        st.exception(exc)
    st.stop()

if df.empty:
    st.warning("No bags found with the selected Raw Name filter. Try a different Raw Name value.")
    st.stop()

file_names = ", ".join(f"**{f.name}**" for f in uploaded_files)
exclusion_note = " (returned/recast bags excluded)" if exclude_recast else ""
st.success(
    f"Processed {df['Bag No'].nunique()} bags from {len(uploaded_files)} file(s): {file_names}{exclusion_note}."
)

duplicate_bags = df["Bag No"][df["Bag No"].duplicated()].unique()
if len(duplicate_bags) > 0:
    st.warning(
        f"⚠️ {len(duplicate_bags)} Bag No appear in more than one uploaded file "
        f"(possible overlapping exports): {', '.join(duplicate_bags[:10])}"
        + (" ..." if len(duplicate_bags) > 10 else "")
    )

categories = sorted([c for c in df["Category"].dropna().unique()])
col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 3])
with col_f1:
    selected_categories = st.multiselect("Filter by Category", categories, default=categories)
with col_f2:
    status_filter = st.multiselect(
        "Filter by Status", ["OK", "Mismatch", "No Data"], default=["OK", "Mismatch", "No Data"]
    )
with col_f3:
    file_options = sorted(df["Source File"].unique())
    selected_files = st.multiselect("Filter by Source File", file_options, default=file_options)
with col_f4:
    search_bag = st.text_input("Search Bag No", placeholder="e.g. 8518")

filtered = df[
    df["Category"].isin(selected_categories)
    & df["Status"].isin(status_filter)
    & df["Source File"].isin(selected_files)
]
if search_bag:
    filtered = filtered[filtered["Bag No"].str.contains(search_bag, case=False, na=False)]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Bags", f"{len(filtered):,}")
k2.metric("Total Opening Wt", f"{filtered['Opening Wt'].sum():,.2f} g")
k3.metric("Total Loss Metal", f"{filtered['Loss Metal'].sum():,.3f} g")
avg_loss_pct = filtered["Gross Loss %"].mean()
k4.metric("Avg Gross Loss %", f"{avg_loss_pct:,.2f} %" if pd.notna(avg_loss_pct) else "-")
mismatch_count = int((filtered["Status"] == "Mismatch").sum())
k5.metric("Mismatches", f"{mismatch_count:,}", delta=None)

tab_table, tab_category, tab_bagcount, tab_department, tab_charts = st.tabs(
    [
        "\U0001F4CB Report",
        "\U0001F5C2️ Category Summary",
        "\U0001F9EE Bags by Category",
        "\U0001F3ED Department Summary",
        "\U0001F4CA Charts",
    ]
)

with tab_table:
    st.dataframe(
        filtered,
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={
            "Gross Loss %": st.column_config.ProgressColumn(
                "Gross Loss %",
                min_value=0,
                max_value=max(float(filtered["Gross Loss %"].max() or 1), 1.0),
                format="%.2f%%",
            ),
            "Status": st.column_config.TextColumn("Status"),
        },
    )

    excel_bytes = to_excel_bytes(filtered)
    st.download_button(
        "⬇️ Download filtered report (.xlsx)",
        data=excel_bytes,
        file_name="Bag Gross Loss % Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

with tab_category:
    weight_summary = build_category_weight_summary(filtered)
    mismatch_summary = build_category_mismatch_summary(filtered)

    st.subheader("Weight Summary by Category")
    st.dataframe(
        weight_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Avg Gross Loss %": st.column_config.NumberColumn("Avg Gross Loss %", format="%.2f%%"),
        },
    )
    st.download_button(
        "⬇️ Download weight summary (.xlsx)",
        data=to_excel_bytes(weight_summary, sheet_name="Category Weight Summary"),
        file_name="Category Weight Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    st.subheader("Mismatch Summary by Category")
    st.caption("Calculated separately from the weight totals above — counts bags by Balance vs Last Weight status.")
    st.dataframe(
        mismatch_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Mismatch %": st.column_config.ProgressColumn(
                "Mismatch %",
                min_value=0,
                max_value=max(float(mismatch_summary["Mismatch %"].max() or 1), 1.0),
                format="%.2f%%",
            ),
        },
    )
    st.download_button(
        "⬇️ Download mismatch summary (.xlsx)",
        data=to_excel_bytes(mismatch_summary, sheet_name="Category Mismatch Summary"),
        file_name="Category Mismatch Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab_bagcount:
    st.subheader("Returned / Recast Bags by Category")
    st.caption(
        "Only bags that went back through CAST a second time with a new In Wt Gm value "
        "(sent back to be reworked/recast, then run through the process again) — not all bags."
    )
    returned_bag_details = build_returned_bag_details(uploaded_files, metal_raw_name=metal_raw_name)
    bag_counts = build_returned_bag_category_counts(returned_bag_details)
    if bag_counts.empty:
        st.info("No returned/recast bags found.")
    else:
        count_col1, count_col2 = st.columns([1, 1])
        with count_col1:
            st.dataframe(
                bag_counts,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "% of Total Bags": st.column_config.ProgressColumn(
                        "% of Total Bags", min_value=0, max_value=100, format="%.2f%%"
                    ),
                },
            )
            st.download_button(
                "⬇️ Download returned/recast bags by category (.xlsx)",
                data=to_excel_bytes(bag_counts, sheet_name="Returned Bags by Category"),
                file_name="Returned Bags by Category.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with count_col2:
            st.bar_chart(bag_counts.set_index("Category")["Returned/Recast Bags"])

        st.divider()
        st.subheader("Bag Numbers")
        cat_options = sorted(returned_bag_details["Category"].dropna().unique())
        cat_filter = st.multiselect(
            "Filter by Category", cat_options, default=cat_options, key="recast_category_filter"
        )
        details_view = returned_bag_details[returned_bag_details["Category"].isin(cat_filter)].sort_values(
            ["Category", "Bag No"]
        )
        st.dataframe(details_view, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download returned/recast bag numbers (.xlsx)",
            data=to_excel_bytes(details_view, sheet_name="Returned Bag Numbers"),
            file_name="Returned Bag Numbers.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_department:
    st.subheader("Summary by Department")
    st.caption(
        "One row per department — same columns as the bag report (Opening Wt, Add Metal, Return Metal, "
        f"Loss Metal, Balance, Gross Loss %), totalled across every '{metal_raw_name}' entry at that department. "
        "Non-metal rows (diamonds/stones) are excluded, same as the rest of the app."
    )
    department_summary = build_department_summary(
        uploaded_files, metal_raw_name=metal_raw_name, exclude_recast=exclude_recast
    )
    st.dataframe(
        department_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Gross Loss %": st.column_config.ProgressColumn(
                "Gross Loss %",
                min_value=0,
                max_value=max(float(department_summary["Gross Loss %"].max() or 1), 1.0),
                format="%.2f%%",
            ),
        },
    )
    st.download_button(
        "⬇️ Download department summary (.xlsx)",
        data=to_excel_bytes(department_summary, sheet_name="Department Summary"),
        file_name="Department Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab_charts:
    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.subheader("Avg Gross Loss % by Category")
        by_cat = filtered.groupby("Category")["Gross Loss %"].mean().sort_values(ascending=False)
        st.bar_chart(by_cat)
    with chart_col2:
        st.subheader("Gross Loss % Distribution")
        hist_df = filtered["Gross Loss %"].dropna()
        if not hist_df.empty:
            bins = pd.cut(hist_df, bins=10)
            counts = bins.value_counts().sort_index()
            counts.index = counts.index.astype(str)
            st.bar_chart(counts)
        else:
            st.caption("No data to chart.")

    st.subheader("Loss Metal by Category")
    loss_by_cat = filtered.groupby("Category")["Loss Metal"].sum().sort_values(ascending=False)
    st.bar_chart(loss_by_cat)
