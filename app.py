import pandas as pd
import dash
from dash import html, dcc, Input, Output
import plotly.express as px

# === Load Excel Data ===
df = pd.read_excel(
    "WRLG_KPI_Daily_Reporting.xlsx",
    sheet_name="KPI 2024 Dump",
    engine="openpyxl",
    header=0
)

# === Format Dates ===
month_map = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}
df["Month"] = df["Month"].map(month_map)
df["Date"] = pd.to_datetime(df[["Year", "Month", "Day"]])
df["MonthStart"] = df["Date"].values.astype("datetime64[M]")
df["MonthYear"] = df["Date"].dt.strftime("%b %Y")

# === Dropdown Options ===
month_lookup = df[["MonthYear", "MonthStart"]].drop_duplicates().sort_values("MonthStart")
month_year_options = month_lookup["MonthYear"].tolist()
material_options = ["Ore", "Waste"]

# === Dash App ===
app = dash.Dash(__name__)
server = app.server

# === Layout ===
app.layout = html.Div([
    html.H1("Underground Mining KPI Dashboard", style={"textAlign": "center"}),

    html.Div([
        html.Label("Select Month(s):"),
        dcc.Dropdown(
            id="month-filter",
            options=[{"label": m, "value": m} for m in month_year_options],
            multi=True,
            value=month_year_options[-1:]
        ),
    ], style={"width": "40%", "margin": "auto"}),

    html.Div(id="kpi-output"),
    html.Div(id="charts-output")
])

# === Callback ===
@app.callback(
    [Output("kpi-output", "children"), Output("charts-output", "children")],
    [Input("month-filter", "value")]
)
def update_dashboard(selected_months):
    filtered = df[df["MonthYear"].isin(selected_months)]

    # KPIs
    total_haul = filtered["Total Hauled Tonnes"].sum()
    ore = filtered["Total Ore Hauled Tonnes"].sum()
    waste = filtered["Total Waste Hauled Tonnes"].sum()
    advance = filtered["Equivalent Advance (m)"].sum()

    kpis = html.Div([
        html.H3(f"Total Hauled: {total_haul:.0f} tonnes"),
        html.H3(f"Ore Hauled: {ore:.0f} tonnes"),
        html.H3(f"Waste Hauled: {waste:.0f} tonnes"),
        html.H3(f"Meters Advanced: {advance:.2f} m")
    ])

    # Charts
    haul_chart = px.bar(
        filtered.groupby("Date")["Total Hauled Tonnes"].sum().reset_index(),
        x="Date", y="Total Hauled Tonnes", title="Total Hauled Tonnes Per Day",
        color=filtered["MonthYear"]
    )

    ore_vs_waste = filtered.groupby("Date")[
        ["Total Ore Hauled Tonnes", "Total Waste Hauled Tonnes"]
    ].sum().reset_index()

    ore_waste_chart = px.bar(
        ore_vs_waste, x="Date",
        y=["Total Ore Hauled Tonnes", "Total Waste Hauled Tonnes"],
        title="Ore vs Waste Hauled"
    )

    advance_chart = px.bar(
        filtered.groupby("Date")["Equivalent Advance (m)"].sum().reset_index(),
        x="Date", y="Equivalent Advance (m)",
        title="Meters Advanced Per Day",
        color=filtered["MonthYear"]
    )

    # Monthly Average Daily Meters Advanced
    daily_avg_by_month = filtered.groupby("MonthYear").agg({
        "Equivalent Advance (m)": "sum",
        "Date": pd.Series.nunique
    }).reset_index()
    daily_avg_by_month["Average Daily Meters"] = daily_avg_by_month["Equivalent Advance (m)"] / daily_avg_by_month["Date"]

    avg_chart = px.bar(
        daily_avg_by_month,
        x="MonthYear", y="Average Daily Meters",
        title="Average Daily Meters Advanced per Month",
        color="MonthYear"
    )

    return kpis, html.Div([
        dcc.Graph(figure=haul_chart),
        dcc.Graph(figure=ore_waste_chart),
        dcc.Graph(figure=advance_chart),
        dcc.Graph(figure=avg_chart)
    ])

# === Run App ===
if __name__ == '__main__':
    app.run(debug=True)
