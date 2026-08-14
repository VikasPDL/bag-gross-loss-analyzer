import pandas as pd
import streamlit as st

from processing import (
    InvalidWorkbookError,
    build_category_department_summary,
    build_category_mismatch_summary,
    build_category_weight_summary,
    build_department_summary,
    build_gross_loss_report,
    build_pure_category_summary,
    build_pure_department_summary,
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
build_category_department_summary = st.cache_data(show_spinner=False)(build_category_department_summary)
build_pure_department_summary = st.cache_data(show_spinner=False)(build_pure_department_summary)

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
file_options = sorted(df["Source File"].unique())

with st.popover("\U0001F50D Filters", use_container_width=False):
    selected_categories = st.multiselect("Filter by Category", categories, default=categories)
    status_filter = st.multiselect(
        "Filter by Status", ["OK", "Mismatch", "No Data"], default=["OK", "Mismatch", "No Data"]
    )
    selected_files = st.multiselect("Filter by Source File", file_options, default=file_options)
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

tab_table, tab_category, tab_bagcount, tab_department, tab_catdept, tab_pure, tab_highloss, tab_charts = st.tabs(
    [
        "\U0001F4CB Report",
        "\U0001F5C2️ Category Summary",
        "\U0000267B\U0000FE0F Recast Bags",
        "\U0001F3ED Department Summary",
        "\U0001F9ED Category & Department",
        "\U0001F48E Pure Weight",
        "\U0001F6A8 High Loss Bags",
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

    cat_rows = weight_summary[weight_summary["Category"] != "TOTAL (by bag count)"]
    top_loss_cat = cat_rows.loc[cat_rows["Avg Gross Loss %"].idxmax()] if not cat_rows.empty else None
    total_mismatch = int(mismatch_summary["Mismatch"].sum())

    oc1, oc2 = st.columns(2)
    with oc1:
        st.markdown(
            stat_card(
                "Category Overview",
                [
                    ("\U0001F5C2\U0000FE0F", "#DCEBFF", "Categories", f"{len(cat_rows):,}"),
                    ("\U0001F4E6", "#FFF1D6", "Total Bags", f"{int(cat_rows['Bags'].sum()):,}"),
                ],
            ),
            unsafe_allow_html=True,
        )
    with oc2:
        st.markdown(
            stat_card(
                "Where It's Worst",
                [
                    (
                        "\U0001F947",
                        "#DFF7EA",
                        "Top Loss Category",
                        f"{top_loss_cat['Category']} ({top_loss_cat['Avg Gross Loss %']:.2f}%)"
                        if top_loss_cat is not None
                        else "—",
                    ),
                    ("\U000026A0\U0000FE0F", "#FFE3E3", "Total Mismatches", f"{total_mismatch:,}"),
                ],
            ),
            unsafe_allow_html=True,
        )
    st.write("")

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
        total_recast = int(bag_counts["Returned/Recast Bags"].sum())
        total_bags_seen = len(filtered) + total_recast if exclude_recast else len(filtered)
        recast_pct = (total_recast / total_bags_seen * 100) if total_bags_seen else 0
        categories_affected = int((bag_counts["Returned/Recast Bags"] > 0).sum())
        top_row = bag_counts.iloc[0]

        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown(
                stat_card(
                    "Recast Overview",
                    [
                        ("\U0000267B\U0000FE0F", "#FFE3E3", "Recast Bags", f"{total_recast:,}"),
                        ("\U0001F4CA", "#FFF1D6", "% of All Bags", f"{recast_pct:,.2f} %"),
                    ],
                ),
                unsafe_allow_html=True,
            )
        with rc2:
            st.markdown(
                stat_card(
                    "Where It's Worst",
                    [
                        ("\U0001F5C2\U0000FE0F", "#DCEBFF", "Categories Affected", f"{categories_affected:,}"),
                        (
                            "\U0001F947",
                            "#DFF7EA",
                            "Top Category",
                            f"{top_row['Category']} ({int(top_row['Returned/Recast Bags'])})",
                        ),
                    ],
                ),
                unsafe_allow_html=True,
            )
        st.write("")

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

        st.divider()
        st.subheader("Recast Bags — Category Summary")
        st.caption("Same table as the Category Summary tab, but scoped to only the recast bags above.")
        recast_bag_report = build_gross_loss_report(
            uploaded_files, metal_raw_name=metal_raw_name, mismatch_tolerance=mismatch_tolerance, only_recast=True
        )
        recast_weight_summary = build_category_weight_summary(recast_bag_report)
        st.dataframe(
            recast_weight_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Avg Gross Loss %": st.column_config.NumberColumn("Avg Gross Loss %", format="%.2f%%"),
            },
        )
        st.download_button(
            "⬇️ Download recast category summary (.xlsx)",
            data=to_excel_bytes(recast_weight_summary, sheet_name="Recast Category Summary"),
            file_name="Recast Category Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.divider()
        st.subheader("Recast Bags — Department Summary")
        st.caption("Same table as the Department Summary tab, but scoped to only the recast bags above.")
        recast_department_summary = build_department_summary(
            uploaded_files, metal_raw_name=metal_raw_name, only_recast=True
        )
        st.dataframe(
            recast_department_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Gross Loss %": st.column_config.ProgressColumn(
                    "Gross Loss %",
                    min_value=0,
                    max_value=max(float(recast_department_summary["Gross Loss %"].max() or 1), 1.0),
                    format="%.2f%%",
                ),
            },
        )
        st.download_button(
            "⬇️ Download recast department summary (.xlsx)",
            data=to_excel_bytes(recast_department_summary, sheet_name="Recast Department Summary"),
            file_name="Recast Department Summary.xlsx",
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

    dept_rows = department_summary[~department_summary["Department"].str.startswith("TOTAL")]
    top_loss_dept = dept_rows.loc[dept_rows["Gross Loss %"].idxmax()] if not dept_rows.empty else None

    reached_fg = df["Last Department"] == "FG"
    fg_count = int(reached_fg.sum())
    fg_pct = (fg_count / len(df) * 100) if len(df) else 0
    fg_final_weight = df.loc[reached_fg, "Last Weight"].sum()

    doc1, doc2, doc3 = st.columns(3)
    with doc1:
        st.markdown(
            stat_card(
                "Department Overview",
                [
                    ("\U0001F3ED", "#DCEBFF", "Departments", f"{len(dept_rows):,}"),
                    ("\U0001F4C9", "#FFE4CC", "Total Loss Metal", f"{dept_rows['Total Loss Metal'].sum():,.3f} g"),
                ],
            ),
            unsafe_allow_html=True,
        )
    with doc2:
        st.markdown(
            stat_card(
                "Where It's Worst",
                [
                    (
                        "\U0001F947",
                        "#DFF7EA",
                        "Top Loss Department",
                        f"{top_loss_dept['Department']} ({top_loss_dept['Gross Loss %']:.2f}%)"
                        if top_loss_dept is not None
                        else "—",
                    ),
                    (
                        "\U0001F4CA",
                        "#FFF1D6",
                        "TOTAL (sum of %)",
                        f"{department_summary['Gross Loss %'].iloc[-1]:.2f} %",
                    ),
                ],
            ),
            unsafe_allow_html=True,
        )
    with doc3:
        st.markdown(
            stat_card(
                "Final Check",
                [
                    ("\U00002705", "#DFF7EA", "Bags Reaching FG", f"{fg_count:,} ({fg_pct:.1f}%)"),
                    ("\U0001F3C1", "#DCEBFF", "Total Final Weight", f"{fg_final_weight:,.2f} g"),
                ],
            ),
            unsafe_allow_html=True,
        )
    st.caption(
        "Final Check uses the Last Weight recorded at FG (Finished Goods) for each bag — the source "
        "system's own final number — rather than the calculated Balance above, since FG has no metal "
        "weighing entry of its own to feed into that chain."
    )
    st.write("")

    fg_row = {col: None for col in department_summary.columns}
    fg_row["Department"] = "FG"
    fg_row["Bags"] = fg_count
    fg_row["Total Opening Wt"] = round(fg_final_weight, 3)
    fg_row["Total Balance"] = round(fg_final_weight, 3)
    department_summary_with_fg = pd.concat(
        [department_summary.iloc[:-1], pd.DataFrame([fg_row]), department_summary.iloc[-1:]],
        ignore_index=True,
    )

    st.dataframe(
        department_summary_with_fg,
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
        data=to_excel_bytes(department_summary_with_fg, sheet_name="Department Summary"),
        file_name="Department Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab_catdept:
    st.subheader("Category & Department, together")
    st.caption("One row per Category + Department, same columns as the bag report.")
    catdept_detail = build_category_department_summary(
        uploaded_files, metal_raw_name=metal_raw_name, exclude_recast=exclude_recast
    )

    if catdept_detail.empty:
        st.info("No Category/Department data found.")
    else:
        catdept_display = catdept_detail[
            ["Category", "Department", "Total Opening Wt", "Total Add Metal", "Total Return Metal",
             "Total Loss Metal", "Total Balance", "Gross Loss %"]
        ].rename(
            columns={
                "Total Opening Wt": "Opening Wt",
                "Total Add Metal": "Add Metal",
                "Total Return Metal": "Return Metal",
                "Total Loss Metal": "Loss Metal",
                "Total Balance": "Balance",
            }
        )

        top_combo = catdept_display.loc[catdept_display["Gross Loss %"].idxmax()]
        cd1, cd2 = st.columns(2)
        with cd1:
            st.markdown(
                stat_card(
                    "Combination Overview",
                    [
                        ("\U0001F9ED", "#DCEBFF", "Combinations", f"{len(catdept_display):,}"),
                        (
                            "\U0001F5C2\U0000FE0F",
                            "#FFF1D6",
                            "Categories",
                            f"{catdept_display['Category'].nunique():,}",
                        ),
                    ],
                ),
                unsafe_allow_html=True,
            )
        with cd2:
            st.markdown(
                stat_card(
                    "Where It's Worst",
                    [
                        (
                            "\U0001F947",
                            "#DFF7EA",
                            "Worst Combination",
                            f"{top_combo['Category']} / {top_combo['Department']} ({top_combo['Gross Loss %']:.2f}%)",
                        ),
                        ("\U0001F3ED", "#FFE3E3", "Departments", f"{catdept_display['Department'].nunique():,}"),
                    ],
                ),
                unsafe_allow_html=True,
            )
        st.write("")

        st.dataframe(
            catdept_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Gross Loss %": st.column_config.ProgressColumn(
                    "Gross Loss %",
                    min_value=0,
                    max_value=max(float(catdept_display["Gross Loss %"].max() or 1), 1.0),
                    format="%.2f%%",
                ),
            },
        )
        st.download_button(
            "⬇️ Download Category & Department (.xlsx)",
            data=to_excel_bytes(catdept_display, sheet_name="Category x Department"),
            file_name="Category x Department.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_pure:
    st.subheader("Pure Weight Overview")
    st.caption(
        "Net weight converted to pure gold weight using each bag's Karat (from Base Metal, e.g. G14K, "
        "G10K) — Pure Wt = Net Wt × Karat / 24. Same columns as everywhere else, just in pure gold terms."
    )

    unknown_karat = filtered["Pure Balance"].isna().sum()
    if unknown_karat:
        st.warning(
            f"⚠️ {unknown_karat} bag(s) have a Base Metal value that couldn't be read as a Karat "
            "(e.g. blank or an unrecognized code) — excluded from the totals below."
        )

    pure_filtered = filtered.dropna(subset=["Pure Balance"])
    pc1, pc2 = st.columns(2)
    with pc1:
        st.markdown(
            stat_card(
                "Pure Weight Overview",
                [
                    ("\U0001F48E", "#DCEBFF", "Total Pure Opening Wt", f"{pure_filtered['Pure Opening Wt'].sum():,.2f} g"),
                    ("\U0001F4C9", "#FFE4CC", "Total Pure Loss Metal", f"{pure_filtered['Pure Loss Metal'].sum():,.3f} g"),
                ],
            ),
            unsafe_allow_html=True,
        )
    with pc2:
        avg_pure_loss = pure_filtered["Pure Gross Loss %"].mean()
        st.markdown(
            stat_card(
                "Performance",
                [
                    (
                        "\U0001F4CA",
                        "#DFF7EA",
                        "Avg Pure Gross Loss %",
                        f"{avg_pure_loss:,.2f} %" if pd.notna(avg_pure_loss) else "—",
                    ),
                    ("\U0001F9EA", "#FFF1D6", "Bags Converted", f"{len(pure_filtered):,}"),
                ],
            ),
            unsafe_allow_html=True,
        )
    st.write("")

    st.subheader("Pure Weight — Category Table")
    pure_category_summary = build_pure_category_summary(pure_filtered)
    st.dataframe(
        pure_category_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Avg Gross Loss %": st.column_config.NumberColumn("Avg Gross Loss %", format="%.2f%%"),
        },
    )
    st.download_button(
        "⬇️ Download Pure Weight category table (.xlsx)",
        data=to_excel_bytes(pure_category_summary, sheet_name="Pure Weight Category"),
        file_name="Pure Weight Category.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    st.subheader("Pure Weight — Department Table")
    pure_department_summary = build_pure_department_summary(
        uploaded_files, metal_raw_name=metal_raw_name, exclude_recast=exclude_recast
    )
    st.dataframe(
        pure_department_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Gross Loss %": st.column_config.ProgressColumn(
                "Gross Loss %",
                min_value=0,
                max_value=max(float(pure_department_summary["Gross Loss %"].max() or 1), 1.0),
                format="%.2f%%",
            ),
        },
    )
    st.download_button(
        "⬇️ Download Pure Weight department table (.xlsx)",
        data=to_excel_bytes(pure_department_summary, sheet_name="Pure Weight Department"),
        file_name="Pure Weight Department.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    st.subheader("Pure Weight — Bag-wise Detail")
    pure_detail = pure_filtered[
        [
            "Source File", "Bag No", "Category", "Karat",
            "Pure Opening Wt", "Pure Add Metal", "Pure Return Metal", "Pure Loss Metal",
            "Pure Balance", "Pure Gross Loss %", "Status",
        ]
    ].rename(
        columns={
            "Pure Opening Wt": "Opening Wt",
            "Pure Add Metal": "Add Metal",
            "Pure Return Metal": "Return Metal",
            "Pure Loss Metal": "Loss Metal",
            "Pure Balance": "Balance",
            "Pure Gross Loss %": "Gross Loss %",
        }
    )
    st.dataframe(
        pure_detail,
        use_container_width=True,
        hide_index=True,
        height=460,
        column_config={
            "Gross Loss %": st.column_config.ProgressColumn(
                "Gross Loss %",
                min_value=0,
                max_value=max(float(pure_detail["Gross Loss %"].max() or 1), 1.0),
                format="%.2f%%",
            ),
        },
    )
    st.download_button(
        "⬇️ Download Pure Weight bag-wise detail (.xlsx)",
        data=to_excel_bytes(pure_detail, sheet_name="Pure Weight Detail"),
        file_name="Pure Weight Detail.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab_highloss:
    st.subheader("Bags With Unusually High Gross Loss %")
    st.caption(
        "Bags above the threshold usually mean the weight chain broke somewhere for that bag "
        "(e.g. Balance shrank to almost nothing, so Loss ÷ Balance blows up) — not a real loss that "
        "large. Shows every uploaded bag regardless of the filters above, so nothing gets missed."
    )
    threshold = st.number_input("Gross Loss % threshold", min_value=0.0, value=100.0, step=10.0)
    high_loss = df[df["Gross Loss %"] > threshold].sort_values("Gross Loss %", ascending=False)

    if high_loss.empty:
        st.success(f"No bags found above {threshold:.0f}% Gross Loss.")
    else:
        st.warning(f"⚠️ {len(high_loss)} bag(s) above {threshold:.0f}% Gross Loss.")

        hl1, hl2 = st.columns(2)
        with hl1:
            st.markdown(
                stat_card(
                    "Outlier Overview",
                    [
                        ("\U0001F6A8", "#FFE3E3", "Bags Above Threshold", f"{len(high_loss):,}"),
                        ("\U0001F4C8", "#FFF1D6", "Highest Gross Loss %", f"{high_loss['Gross Loss %'].max():,.1f} %"),
                    ],
                ),
                unsafe_allow_html=True,
            )
        with hl2:
            st.markdown(
                stat_card(
                    "Where It's Worst",
                    [
                        (
                            "\U0001F5C2\U0000FE0F",
                            "#DCEBFF",
                            "Top Category",
                            high_loss["Category"].mode().iat[0],
                        ),
                        ("\U00002139\U0000FE0F", "#DFF7EA", "Most Common Status", high_loss["Status"].mode().iat[0]),
                    ],
                ),
                unsafe_allow_html=True,
            )
        st.write("")

        st.dataframe(
            high_loss,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Gross Loss %": st.column_config.NumberColumn("Gross Loss %", format="%.1f%%"),
            },
        )
        st.download_button(
            "⬇️ Download high loss bags (.xlsx)",
            data=to_excel_bytes(high_loss, sheet_name="High Loss Bags"),
            file_name="High Loss Bags.xlsx",
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
