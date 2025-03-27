import pandas as pd
import dash
from dash import html, dcc, Output, Input
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

app.layout = html.Div([
    html.H1("Underground Mining KPI Dashboard", style={"textAlign": "center"}),

    html.Div([
        html.Div([
            html.Label("Select Month(s):"),
            dcc.Dropdown(
                id="month-dropdown",
                options=[{"label": m, "value": m} for m in month_year_options],
                value=[month_year_options[-1]],
                multi=True
            )
        ], style={"width": "48%", "display": "inline-block", "padding": "10px"}),

        html.Div([
            html.Label("Select Material Type(s):"),
            dcc.Dropdown(
                id="material-dropdown",
                options=[{"label": m, "value": m} for m in material_options],
                value=material_options,
                multi=True
            )
        ], style={"width": "48%", "display": "inline-block", "padding": "10px"})
    ]),

    html.Div(id="kpi-output", style={"display": "flex", "flexWrap": "wrap", "gap": "20px", "padding": "20px"}),
    html.Div(id="charts-output", style={"padding": "20px"})
])

# === Callback ===
@app.callback(
    [Output("kpi-output", "children"),
     Output("charts-output", "children")],
    [Input("month-dropdown", "value"),
     Input("material-dropdown", "value")]
)
def update_dashboard(selected_months, selected_materials):
    filtered = df[df["MonthYear"].isin(selected_months)]

    if filtered.empty:
        return html.Div("⚠️ No data available for the selected filters."), html.Div()

    # Prepare grouped data for charts
    grouped = filtered.copy()
    grouped["MonthLabel"] = grouped["Date"].dt.strftime("%b %Y")
    grouped_daily = grouped.groupby(["Date", "MonthLabel"]).sum(numeric_only=True).reset_index()

    # NEW: Proper average daily meters per month
    daily_totals = grouped.groupby(["Date", "MonthLabel"])["Equivalent Advance (m)"].sum().reset_index()
    monthly_avg = daily_totals.groupby("MonthLabel")["Equivalent Advance (m)"].mean().reset_index()

    # Rolling average for Meters Advanced
    rolling = grouped_daily[["Date", "Equivalent Advance (m)"]].copy()
    rolling["Rolling 30-Day Avg"] = rolling["Equivalent Advance (m)"].rolling(window=30).mean()

    # KPI totals
    show_ore = "Ore" in selected_materials
    show_waste = "Waste" in selected_materials

    total_haul = 0
    total_ore = 0
    total_waste = 0
    total_advance = grouped["Equivalent Advance (m)"].sum()

    if show_ore:
        total_ore = grouped["Total Ore Hauled Tonnes"].sum()
        total_haul += total_ore
    if show_waste:
        total_waste = grouped["Total Waste Hauled Tonnes"].sum()
        total_haul += total_waste

    # === KPI Cards ===
    def kpi_card(label, value):
        return html.Div([
            html.H3(label),
            html.P(f"{value:,}" if isinstance(value, (int, float)) else value)
        ], style={
            "border": "1px solid #ccc",
            "borderRadius": "10px",
            "padding": "15px",
            "minWidth": "200px",
            "textAlign": "center",
            "boxShadow": "2px 2px 10px #eee"
        })

    kpis = [
        kpi_card("📅 Selected Months", ", ".join(selected_months)),
        kpi_card("🪨 Selected Materials", ", ".join(selected_materials)),
        kpi_card("🚛 Total Hauled Tonnes", total_haul),
        kpi_card("🟢 Ore Hauled", total_ore),
        kpi_card("🟤 Waste Hauled", total_waste),
        kpi_card("📏 Equivalent Advance (m)", total_advance)
    ]

    # === Charts ===

    # Total Hauled Tonnes Chart
    haul_chart = px.bar(
        grouped_daily,
        x="Date",
        y="Total Hauled Tonnes",
        color="MonthLabel",
        title="Total Hauled Tonnes by Date and Month",
        labels={"Date": "Date", "Total Hauled Tonnes": "Tonnes", "MonthLabel": "Month"}
    )

    # Ore vs Waste Chart
    material_chart_data = []
    if show_ore:
        material_chart_data.append(("Ore Hauled", total_ore))
    if show_waste:
        material_chart_data.append(("Waste Hauled", total_waste))

    ore_vs_waste_chart = px.bar(
        x=[label for label, _ in material_chart_data],
        y=[value for _, value in material_chart_data],
        labels={"x": "Material", "y": "Tonnes"},
        title="Ore vs Waste Hauled"
    )

    # Meters Advanced Chart with Rolling Average
    advance_fig = px.bar(
        grouped_daily,
        x="Date",
        y="Equivalent Advance (m)",
        color="MonthLabel",
        title="Meters Advanced (with Rolling Average)",
        labels={"Date": "Date", "Equivalent Advance (m)": "Meters", "MonthLabel": "Month"}
    )
    advance_fig.add_scatter(
        x=rolling["Date"],
        y=rolling["Rolling 30-Day Avg"],
        mode="lines",
        name="30-Day Rolling Avg",
        line=dict(color="black", dash="dot")
    )

    # Average Daily Meters per Month (New)
    monthly_avg_chart = px.bar(
        monthly_avg,
        x="MonthLabel",
        y="Equivalent Advance (m)",
        color="MonthLabel",
        labels={"MonthLabel": "Month", "Equivalent Advance (m)": "Avg Daily Meters"},
        title="Average Daily Meters Advanced per Month"
    )

    charts = html.Div([
        dcc.Graph(figure=haul_chart),
        dcc.Graph(figure=ore_vs_waste_chart),
        dcc.Graph(figure=advance_fig),
        dcc.Graph(figure=monthly_avg_chart)
    ])

    return kpis, charts

# === Run App ===
if __name__ == '__main__':
    app.run(debug=True)
