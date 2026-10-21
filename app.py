from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Transport for Clients Assessment",
    page_icon="🚐",
    layout="wide",
)


DEFAULT_WORKBOOK = "Transport_for_Clients (1).xlsx"


PROVIDER_NAME_CORRECTIONS = {
    "SLEC Transpor Office (July - Aug 2026 Pilot)":
        "SLEC Transport Office (July - Aug 2026 Pilot)",
}


def clean_provider_name(provider):
    """Return a consistent provider name for matching across worksheets."""
    provider = str(provider).strip()
    return PROVIDER_NAME_CORRECTIONS.get(provider, provider)


def read_measure_sheet(file_source, sheet_name, value_name):
    """Convert one wide worksheet into provider-month rows."""
    wide_data = pd.read_excel(file_source, sheet_name=sheet_name, header=None)

    month_values = pd.to_datetime(wide_data.iloc[0, 1:], errors="coerce")
    valid_columns = [
        column_number + 1
        for column_number, month in enumerate(month_values)
        if pd.notna(month)
    ]

    if not valid_columns:
        raise ValueError(f"No month columns were found in the '{sheet_name}' sheet.")

    selected_data = wide_data.iloc[1:, [0] + valid_columns].copy()
    selected_data.columns = [
        "Provider",
        *[pd.Timestamp(wide_data.iloc[0, column]) for column in valid_columns],
    ]

    selected_data = selected_data.dropna(subset=["Provider"])
    selected_data["Provider"] = selected_data["Provider"].apply(
        clean_provider_name
    )

    long_data = selected_data.melt(
        id_vars="Provider",
        var_name="Month",
        value_name=value_name,
    )

    long_data["Month"] = pd.to_datetime(long_data["Month"])
    long_data[value_name] = pd.to_numeric(
        long_data[value_name], errors="coerce"
    ).fillna(0)

    return long_data


@st.cache_data
def load_transport_data(file_source):
    """Read trip and cost sheets and combine them into one dataset."""
    trips = read_measure_sheet(file_source, "trips", "Trips")
    costs = read_measure_sheet(file_source, "cost", "Cost")

    combined = pd.merge(
        trips,
        costs,
        on=["Provider", "Month"],
        how="outer",
    )

    combined["Trips"] = combined["Trips"].fillna(0)
    combined["Cost"] = combined["Cost"].fillna(0)
    combined["Trips"] = combined["Trips"].round().astype(int)

    combined["Average cost per trip"] = 0.0
    has_trips = combined["Trips"] > 0
    combined.loc[has_trips, "Average cost per trip"] = (
        combined.loc[has_trips, "Cost"]
        / combined.loc[has_trips, "Trips"]
    )

    return combined.sort_values(["Month", "Provider"]).reset_index(drop=True)


def find_default_workbook():
    """Look for the workbook beside the app and in the current folder."""
    possible_locations = [
        Path(__file__).resolve().parent / DEFAULT_WORKBOOK,
        Path.cwd() / DEFAULT_WORKBOOK,
        Path.cwd() / "Transport_for_Clients.xlsx",
    ]

    for location in possible_locations:
        if location.exists():
            return location

    return None


def format_period(start_month, end_month):
    """Create the period label used throughout the output."""
    start_text = pd.Timestamp(start_month).strftime("%b %Y")
    end_text = pd.Timestamp(end_month).strftime("%b %Y")
    number_of_months = (
        (pd.Timestamp(end_month).year - pd.Timestamp(start_month).year) * 12
        + pd.Timestamp(end_month).month
        - pd.Timestamp(start_month).month
        + 1
    )

    if start_month == end_month:
        return f"{start_text} (1 month)", number_of_months

    return f"{start_text} to {end_text} ({number_of_months} months)", number_of_months


def build_provider_summary(filtered_data):
    """Summarise trips, cost and unit cost by provider."""
    summary = (
        filtered_data.groupby("Provider", as_index=False)
        .agg(
            Trips=("Trips", "sum"),
            Cost=("Cost", "sum"),
        )
    )

    summary["Average cost per trip"] = 0.0
    has_trips = summary["Trips"] > 0
    summary.loc[has_trips, "Average cost per trip"] = (
        summary.loc[has_trips, "Cost"]
        / summary.loc[has_trips, "Trips"]
    )

    total_cost = summary["Cost"].sum()
    summary["Cost share"] = 0.0
    if total_cost > 0:
        summary["Cost share"] = summary["Cost"] / total_cost

    return summary.sort_values("Cost", ascending=False).reset_index(drop=True)


def create_download_data(filtered_data, period_label):
    """Prepare a detailed CSV with the assessment period included."""
    download_data = filtered_data.copy()
    download_data.insert(0, "Assessment period", period_label)
    download_data["Month"] = download_data["Month"].dt.strftime("%Y-%m")
    return download_data.to_csv(index=False).encode("utf-8-sig")


def main():
    st.title("Transport for Clients Assessment")
    st.caption(
        "Review trip volume, expenditure and average cost across transport providers."
    )

    default_workbook = find_default_workbook()

    with st.sidebar:
        st.header("Data and filters")
        uploaded_workbook = st.file_uploader(
            "Upload transport workbook",
            type=["xlsx"],
            help="The workbook must contain sheets named 'trips' and 'cost'.",
        )

    file_source = uploaded_workbook or default_workbook

    if file_source is None:
        st.info(
            f"Upload the workbook, or place '{DEFAULT_WORKBOOK}' beside app.py."
        )
        st.stop()

    try:
        transport_data = load_transport_data(file_source)
    except Exception as error:
        st.error(f"The workbook could not be read: {error}")
        st.stop()

    available_months = sorted(transport_data["Month"].dropna().unique())
    available_providers = sorted(transport_data["Provider"].dropna().unique())

    with st.sidebar:
        selected_providers = st.multiselect(
            "Providers",
            options=available_providers,
            default=available_providers,
        )

        start_month = st.selectbox(
            "Assessment start month",
            options=available_months,
            index=0,
            format_func=lambda month: pd.Timestamp(month).strftime("%b %Y"),
        )

        valid_end_months = [
            month for month in available_months if month >= start_month
        ]
        end_month = st.selectbox(
            "Assessment end month",
            options=valid_end_months,
            index=len(valid_end_months) - 1,
            format_func=lambda month: pd.Timestamp(month).strftime("%b %Y"),
        )

    if not selected_providers:
        st.warning("Select at least one provider to continue.")
        st.stop()

    filtered_data = transport_data[
        transport_data["Provider"].isin(selected_providers)
        & (transport_data["Month"] >= start_month)
        & (transport_data["Month"] <= end_month)
    ].copy()

    period_label, number_of_months = format_period(start_month, end_month)
    provider_summary = build_provider_summary(filtered_data)

    total_trips = int(filtered_data["Trips"].sum())
    total_cost = float(filtered_data["Cost"].sum())
    average_cost = total_cost / total_trips if total_trips else 0
    active_providers = int(
        provider_summary.loc[
            (provider_summary["Trips"] > 0) | (provider_summary["Cost"] > 0),
            "Provider",
        ].nunique()
    )

    st.subheader("Assessment period")
    st.markdown(f"### {period_label}")
    st.caption(
        f"Selected providers: {len(selected_providers)} | "
        f"Providers with recorded activity: {active_providers}"
    )

    metric_1, metric_2, metric_3, metric_4 = st.columns(4)
    metric_1.metric("Total trips", f"{total_trips:,}")
    metric_2.metric("Total cost", f"S${total_cost:,.2f}")
    metric_3.metric("Average cost per trip", f"S${average_cost:,.2f}")
    metric_4.metric("Average monthly cost", f"S${total_cost / number_of_months:,.2f}")

    monthly_summary = (
        filtered_data.groupby("Month", as_index=False)
        .agg(Trips=("Trips", "sum"), Cost=("Cost", "sum"))
        .sort_values("Month")
    )

    chart_1, chart_2 = st.columns(2)

    with chart_1:
        st.markdown(f"#### Monthly trips: {period_label}")
        st.line_chart(
            monthly_summary.set_index("Month")[["Trips"]],
            use_container_width=True,
        )

    with chart_2:
        st.markdown(f"#### Monthly cost: {period_label}")
        st.line_chart(
            monthly_summary.set_index("Month")[["Cost"]],
            use_container_width=True,
        )

    st.subheader("Provider assessment")
    display_summary = provider_summary.copy()
    display_summary["Cost"] = display_summary["Cost"].map(
        lambda value: f"S${value:,.2f}"
    )
    display_summary["Average cost per trip"] = display_summary[
        "Average cost per trip"
    ].map(lambda value: f"S${value:,.2f}")
    display_summary["Cost share"] = display_summary["Cost share"].map(
        lambda value: f"{value:.1%}"
    )
    st.dataframe(display_summary, use_container_width=True, hide_index=True)

    provider_chart_data = provider_summary[provider_summary["Cost"] > 0]
    if not provider_chart_data.empty:
        st.markdown(f"#### Cost by provider: {period_label}")
        st.bar_chart(
            provider_chart_data.set_index("Provider")[["Cost"]],
            use_container_width=True,
        )

    st.subheader("Assessment output")
    st.write(
        f"For **{period_label}**, the selected providers recorded "
        f"**{total_trips:,} trips** at a total cost of "
        f"**S${total_cost:,.2f}**, equivalent to an average of "
        f"**S${average_cost:,.2f} per trip**."
    )

    csv_data = create_download_data(filtered_data, period_label)
    file_period = (
        f"{pd.Timestamp(start_month).strftime('%Y%m')}_"
        f"{pd.Timestamp(end_month).strftime('%Y%m')}"
    )
    st.download_button(
        "Download filtered assessment",
        data=csv_data,
        file_name=f"transport_assessment_{file_period}.csv",
        mime="text/csv",
    )

    with st.expander("View monthly provider data"):
        detailed_display = filtered_data.copy()
        detailed_display["Month"] = detailed_display["Month"].dt.strftime(
            "%b %Y"
        )
        st.dataframe(detailed_display, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
