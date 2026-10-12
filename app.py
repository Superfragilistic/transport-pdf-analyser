import io
import os
import datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title="Transport Data Analyzer",
    page_icon="📊",
    layout="centered"
)

# --- Custom CSS Styling ---
st.markdown("""
    <style>
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #1A365D;
        text-align: center;
        margin-bottom: 5px;
    }
    .subtitle {
        font-size: 16px;
        color: #4A5568;
        text-align: center;
        margin-bottom: 30px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Transport Data Analyzer Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your Excel template to compile a comprehensive PDF Analysis Report instantly.</div>', unsafe_allow_html=True)

# --- Core Data Processing & PDF Generator Function ---
def process_and_generate_pdf(file_contents, output_pdf_name='Transport_Analysis_Summary.pdf'):
    # 1. Load data from the uploaded file bytes
    excel_file = io.BytesIO(file_contents)
    df_sheet1 = pd.read_excel(excel_file, sheet_name=0)
    df_sheet2 = pd.read_excel(excel_file, sheet_name=1)

    # 2. Rename columns
    df_sheet1.rename(columns={df_sheet1.columns[0]: 'Provider'}, inplace=True)
    df_sheet2.rename(columns={df_sheet2.columns[0]: 'Provider'}, inplace=True)

    # 3. Format Date Columns to 'MMM YYYY'
    def format_columns(df):
        new_columns = {}
        for col in df.columns:
            if isinstance(col, (datetime.datetime, pd.Timestamp)):
                new_columns[col] = col.strftime('%b %Y')
            elif isinstance(col, str):
                try:
                    parsed_date = pd.to_datetime(col)
                    new_columns[col] = parsed_date.strftime('%b %Y')
                except (ValueError, TypeError):
                    pass
        return df.rename(columns=new_columns)

    df_sheet1 = format_columns(df_sheet1)
    df_sheet2 = format_columns(df_sheet2)

    # 4. Drop Unnamed/Helper Columns
    cols_to_drop = [col for col in df_sheet1.columns if 'Unnamed' in str(col)]
    df_sheet1.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    df_sheet2.drop(columns=cols_to_drop, inplace=True, errors='ignore')

    # Ensure Provider names are clean strings and filter out numeric or empty rows
    df_sheet1['Provider'] = df_sheet1['Provider'].fillna("").astype(str).str.strip()
    df_sheet2['Provider'] = df_sheet2['Provider'].fillna("").astype(str).str.strip()

    # Filter out empty, total, spacer or numeric dummy rows
    invalid_names = ['', '0', '0.0', 'nan', 'nan.0', 'Total', 'Totals', 'total', 'totals']
    df_sheet1 = df_sheet1[~df_sheet1['Provider'].isin(invalid_names)].reset_index(drop=True)
    df_sheet2 = df_sheet2[~df_sheet2['Provider'].isin(invalid_names)].reset_index(drop=True)

    # Fill NaNs with zeros in data columns
    df_sheet1.fillna(0, inplace=True)
    df_sheet2.fillna(0, inplace=True)

    # Drop specific structural empty/totals spacer rows if they exist at index 6 and 7
    df_sheet1_cleaned = df_sheet1.drop(index=[6, 7], errors='ignore').reset_index(drop=True)
    df_sheet2_cleaned = df_sheet2.drop(index=[6, 7], errors='ignore').reset_index(drop=True)

    # 7. Perform Calculations
    month_cols = [col for col in df_sheet1_cleaned.columns if col != 'Provider']

    df_trips_summary = pd.DataFrame({
        'Provider': df_sheet1_cleaned['Provider'],
        'Total Trips': df_sheet1_cleaned[month_cols].sum(axis=1)
    })
    df_trips_summary['Ave_trips'] = df_trips_summary['Total Trips'] / 12

    df_cost_summary = pd.DataFrame({
        'Provider': df_sheet2_cleaned['Provider'],
        'Total Cost': df_sheet2_cleaned[month_cols].sum(axis=1)
    })
    df_cost_summary['Ave_cost_per_month'] = df_cost_summary['Total Cost'] / 12

    df_summary_merged = pd.merge(df_trips_summary, df_cost_summary, on='Provider')
    df_summary_merged['Ave_cost_per_trip'] = df_summary_merged['Total Cost'] / df_summary_merged['Total Trips']
    df_summary_merged.fillna(0, inplace=True)

    # 8. Plot and Save Line Chart
    df_trips_long_clean = df_sheet1_cleaned.melt(
        id_vars=['Provider'],
        var_name='Month', 
        value_name='Trips'
    )

    plt.figure(figsize=(10, 5))
    sns.set_theme(style="whitegrid")
    sns.lineplot(
        data=df_trips_long_clean,
        x='Month',
        y='Trips',
        hue='Provider',
        marker='o',
        linewidth=2
    )
    plt.title('Comparison of Monthly Trip Trends by Provider', fontsize=12, fontweight='bold')
    plt.xlabel('Month', fontsize=10)
    plt.ylabel('Number of Trips', fontsize=10)
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.legend(title='Providers', bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=8)
    plt.tight_layout()

    chart_path = 'temp_line_chart.png'
    plt.savefig(chart_path, dpi=300)
    plt.close()

    # 9. Build ReportLab PDF
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#1A365D'),
        spaceAfter=15
    )
    section_style = ParagraphStyle(
        'SectionStyle',
        parent=styles['Heading2'],
        fontSize=14,
        leading=18,
        textColor=colors.HexColor('#2C3E50'),
        spaceAfter=10
    )
    normal_style = styles['Normal']

    # Header components
    story.append(Paragraph("Transport Usage & Costs Analysis Report", title_style))
    story.append(Paragraph("Summary of client trips, monthly averages, and operational expenses generated via Web Upload.", normal_style))
    story.append(Spacer(1, 15))

    # Trend Chart Section
    story.append(Paragraph("1. Monthly Trip Trends Comparison", section_style))
    story.append(Image(chart_path, width=500, height=250))
    story.append(Spacer(1, 20))

    # Summary Table Section
    story.append(Paragraph("2. Consolidated Provider Summary Table", section_style))

    df_formatted = df_summary_merged.round(2).copy()
    table_data = [[Paragraph(f"<b>{col}</b>", styles['Normal']) for col in df_formatted.columns]]

    for _, row in df_formatted.iterrows():
        row_cells = []
        for i, val in enumerate(row):
            if i == 0:
                row_cells.append(Paragraph(str(val), styles['Normal']))
            elif 'Cost' in df_formatted.columns[i] or 'cost' in df_formatted.columns[i]:
                row_cells.append(f"${val:,.2f}")
            else:
                row_cells.append(f"{val:,.2f}")
        table_data.append(row_cells)

    summary_table = Table(table_data, colWidths=[180, 60, 60, 70, 70, 80])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F4F4')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#1A365D')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D5D8DC')),
    ]))

    story.append(summary_table)
    doc.build(story)

    # Clean up temp files
    if os.path.exists(chart_path):
        os.remove(chart_path)

    buffer.seek(0)
    return buffer, df_summary_merged, df_trips_long_clean

# --- 10. Streamlit File Upload and Run Interface ---
uploaded_file = st.file_uploader("Choose your input Transport Excel file:", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded successfully! Clicking below will generate your customized PDF.")

    # Read file contents once
    file_bytes = uploaded_file.read()

    # Generate outputs
    try:
        pdf_buffer, df_summary, df_long = process_and_generate_pdf(file_bytes)

        # Download Button
        st.download_button(
            label="📥 Download PDF Summary Report",
            data=pdf_buffer,
            file_name="Transport_Analysis_Summary.pdf",
            mime="application/pdf"
        )

        # Render Live Previews
        st.subheader("Live App Data Preview")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Consolidated Averages:**")
            st.dataframe(df_summary.style.format(precision=2))
        with col2:
            # Display the trend visualization inside the app state
            st.markdown("**Monthly Client Trip Trends:**")
            fig, ax = plt.subplots(figsize=(6, 3.5))
            sns.lineplot(data=df_long, x='Month', y='Trips', hue='Provider', marker='o', linewidth=2, ax=ax)
            plt.xticks(rotation=45, ha='right', fontsize=6)
            plt.yticks(fontsize=6)
            ax.legend(prop={'size': 5})
            st.pyplot(fig)

    except Exception as e:
        st.error(f"Error processing file: {e}")
