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

def stat_card(title, items):
    stat_html = "".join(
        f'''
        <div style="display:flex; align-items:center; gap:12px; flex:1; min-width:150px;">
            <div style="width:44px;height:44px;border-radius:12px;background:{color};
                        display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;">{icon}</div>
            <div>
                <div style="font-size:0.76rem;color:#8A93A6;font-weight:600;">{label}</div>
                <div style="font-size:1.3rem;font-weight:700;color:#1B2430;">{value}</div>
            </div>
        </div>'''
        for icon, color, label, value in items
    )
    return f'''
    <div style="background:#FFFFFF;border-radius:16px;padding:18px 20px;
                box-shadow:0 1px 3px rgba(16,24,40,0.06);border:1px solid rgba(15,23,42,0.06);height:100%;">
        <div style="font-size:0.92rem;font-weight:700;color:#1B2430;margin-bottom:14px;">{title}</div>
        <div style="display:flex; gap:20px; flex-wrap:wrap;">{stat_html}</div>
    </div>'''


build_gross_loss_report = st.cache_data(show_spinner=False)(build_gross_loss_report)
build_department_summary = st.cache_data(show_spinner=False)(build_department_summary)
build_returned_bag_details = st.cache_data(show_spinner=False)(build_returned_bag_details)

st.set_page_config(
    page_title="Bag Gross Loss % Analyzer",
    page_icon="\U0001F48D",
    layout="wide",
)

PAGE_BG = "#E9F1FB"

CUSTOM_CSS = f"""
<style>
#MainMenu, footer {{visibility: hidden;}}

[data-testid="stAppViewContainer"] {{
    background: {PAGE_BG};
}}
[data-testid="stHeader"] {{
    background: transparent;
}}

.block-container {{
    padding-top: 1.75rem;
    padding-bottom: 2.5rem;
    max-width: 1400px;
    background: #FFFFFF;
    border-radius: 24px;
    margin-top: 1rem;
    margin-bottom: 1rem;
    box-shadow: 0 8px 24px rgba(31, 41, 55, 0.08);
}}

[data-testid="stSidebar"] {{
    background: #FFFFFF;
    border-radius: 0 24px 24px 0;
    margin: 1rem 0 1rem 1rem;
    box-shadow: 0 8px 24px rgba(31, 41, 55, 0.08);
}}
[data-testid="stSidebar"] > div:first-child {{
    padding-top: 1.5rem;
}}

/* Logo header */
.app-logo-row {{
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 4px;
}}
.app-logo-badge {{
    width: 46px;
    height: 46px;
    border-radius: 13px;
    background: linear-gradient(135deg, #2FB380, #1E9468);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 22px;
    flex-shrink: 0;
}}
.app-logo-title {{
    font-size: 1.55rem;
    font-weight: 700;
    color: #1B2430;
    letter-spacing: -0.01em;
}}

[data-testid="stCaptionContainer"] {{
    font-size: 0.95rem !important;
}}

/* Section headers (st.subheader) */
h3 {{
    font-size: 1.2rem !important;
    font-weight: 600 !important;
    margin-top: 0.5rem !important;
}}

/* Sidebar section labels */
[data-testid="stSidebar"] h2 {{
    font-size: 0.85rem !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: #6B7280 !important;
}}

/* Tabs — pill style */
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    border-bottom: none;
    background: #F1F4F9;
    border-radius: 12px;
    padding: 5px;
}}
.stTabs [data-baseweb="tab"] {{
    height: 42px;
    padding: 0 18px;
    border-radius: 9px;
}}
.stTabs [data-baseweb="tab"] p {{
    font-size: 0.92rem !important;
    font-weight: 600 !important;
    color: #6B7280;
}}
.stTabs [aria-selected="true"] {{
    background-color: #FFFFFF;
    box-shadow: 0 1px 4px rgba(31, 41, 55, 0.12);
}}
.stTabs [aria-selected="true"] p {{
    color: #1B2430 !important;
}}

/* Buttons */
.stDownloadButton button {{
    border-radius: 999px;
    font-weight: 600;
    background: #2FB380;
    color: white;
    border: none;
}}
.stButton button {{
    border-radius: 999px;
    font-weight: 600;
}}

/* Multiselect pills */
[data-baseweb="tag"] {{
    border-radius: 999px !important;
    background-color: #E6F6EF !important;
}}
[data-baseweb="tag"] span {{
    color: #1E9468 !important;
}}

/* Labels for inputs/filters */
[data-testid="stWidgetLabel"] p {{
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    color: #4B5563 !important;
}}

hr {{margin: 1.25rem 0 !important;}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.markdown(
    '<div class="app-logo-row"><div class="app-logo-badge">\U0001F48D</div>'
    '<div class="app-logo-title">Bag Gross Loss % Analyzer</div></div>',
    unsafe_allow_html=True,
)
st.caption(
    "Upload the Bag Transaction export to compute Opening / Add Metal / Return Metal / "
    "Loss Metal / Balance and Gross Loss % for every bag, one row per bag."
)

with st.sidebar:
    st.header("\U0001F4E4 Upload")
    uploaded_files = st.file_uploader(
        "Bag Transaction Excel file(s) (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more series exports (e.g. 25, 26) — they'll be combined into one report.",
    )
    st.divider()
    with st.expander("\U00002699\U0000FE0F Settings"):
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

avg_loss_pct = filtered["Gross Loss %"].mean()
mismatch_count = int((filtered["Status"] == "Mismatch").sum())

kc1, kc2, kc3 = st.columns(3)
with kc1:
    st.markdown(
        stat_card(
            "Bag Overview",
            [
                ("\U0001F4E6", "#DCEBFF", "Bags", f"{len(filtered):,}"),
                ("\U000026A0\U0000FE0F", "#FFE3E3", "Mismatches", f"{mismatch_count:,}"),
            ],
        ),
        unsafe_allow_html=True,
    )
with kc2:
    st.markdown(
        stat_card(
            "Weight Overview",
            [
                ("\U00002696\U0000FE0F", "#FFF1D6", "Total Opening Wt", f"{filtered['Opening Wt'].sum():,.2f} g"),
                ("\U0001F4C9", "#FFE4CC", "Total Loss Metal", f"{filtered['Loss Metal'].sum():,.3f} g"),
            ],
        ),
        unsafe_allow_html=True,
    )
with kc3:
    st.markdown(
        stat_card(
            "Performance",
            [
                (
                    "\U0001F4CA",
                    "#DFF7EA",
                    "Avg Gross Loss %",
                    f"{avg_loss_pct:,.2f} %" if pd.notna(avg_loss_pct) else "-",
                ),
            ],
        ),
        unsafe_allow_html=True,
    )
st.write("")

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
            st.bar_chart(bag_counts.set_index("Category")["Returned/Recast Bags"], color="#2FB380")

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
        st.bar_chart(by_cat, color="#2FB380")
    with chart_col2:
        st.subheader("Gross Loss % Distribution")
        hist_df = filtered["Gross Loss %"].dropna()
        if not hist_df.empty:
            bins = pd.cut(hist_df, bins=10)
            counts = bins.value_counts().sort_index()
            counts.index = counts.index.astype(str)
            st.bar_chart(counts, color="#2FB380")
        else:
            st.caption("No data to chart.")

    st.subheader("Loss Metal by Category")
    loss_by_cat = filtered.groupby("Category")["Loss Metal"].sum().sort_values(ascending=False)
    st.bar_chart(loss_by_cat, color="#2FB380")
