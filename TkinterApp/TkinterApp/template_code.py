
import pandas as pd
import json
from datetime import datetime
#import dash
#from dash import dcc, html, Input, Output, State
import plotly.graph_objs as go


# --- Источник данных ---
DATA_FILE_PATH = "your_data.xlsx"  # ← заменить на путь к файлу
DATE_COLUMN = "Дата ДТП"
WEEK_COLUMN = "Неделя"
YEAR_COLUMN = "Год"

# --- Столбцы для фильтров и дашбордов ---
COLUMNS = {
    "week": WEEK_COLUMN,
    "year": YEAR_COLUMN,
    "district": "Район",
    "cause": "Причина ДТП",
    "branch": "Название филиала",
}

# --- Настройки фильтров ---
FILTER_CONFIG = {
    "week": {"type": "range", "width_pct": 30},
    "cause": {"width_pct": 30},
    "district": {"width_pct": None},  # стандартная ширина
}

# --- Дашборды ---
DASHBOARDS = [
    {
        "title": "Top Causes",
        "data_column": COLUMNS["cause"],
        "top_n": 10,
        "color": "#C73E1D",
        "width_pct": 30,
    },
    {
        "title": "Top Districts",
        "data_column": COLUMNS["district"],
        "top_n": 10,
        "color": "#3BB273",
        "width_pct": 30,
    },
    {
        "title": "Top Branches",
        "data_column": COLUMNS["branch"],
        "top_n": 10,
        "color": "#4A90E2",
        "width_pct": None,  # будет 45% по умолчанию
    },
]

# --- Глобальные настройки ---
LAYOUT = {
    "default_width_pct": 45,
    "dashboards_per_row": 2,
    "exclude_nan": True,
}

# =============================
# 2. ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ
# =============================

df = pd.read_excel(DATA_FILE_PATH)

# Создание колонки "Неделя года" как строка: "1 нед 2025 года"
df["WeekLabel"] = df[COLUMNS["week"]].astype(str) + " нед " + df[COLUMNS["year"]].astype(str) + " года"

# Уникальные недели в порядке возрастания
weeks_list = sorted(df["WeekLabel"].unique(), key=lambda x: (
    int(x.split()[0]), int(x.split()[-2])
))

# =============================
# 3. ИНИЦИАЛИЗАЦИЯ DASH-ПРИЛОЖЕНИЯ
# =============================

app = dash.Dash(__name__)

# =============================
# 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================

def aggregate_top(data, column, top_n=10, exclude_nan=True):
    if exclude_nan:
        data = data.dropna(subset=[column])
    counts = data[column].value_counts().head(top_n)
    return counts.index.tolist(), counts.values.tolist()

def create_bar_chart(x, y, title, color, height=180, width=510):
    return go.Figure(
        data=[go.Bar(x=x, y=y, marker_color=color)],
        layout=go.Layout(
            title=title,
            margin=dict(t=25, b=60, l=45, r=15),
            xaxis=dict(tickangle=-45),
            height=height,
            width=width,
        )
    )

# =============================
# 5. LAYOUT
# =============================

# Фильтры
filter_elements = []

# Week range filter
filter_elements.append(html.Div([
    html.Label("Неделя (диапазон):", style={"fontWeight": "bold"}),
    dcc.RangeSlider(
        id="week-range",
        min=0,
        max=len(weeks_list) - 1,
        step=1,
        value=[0, len(weeks_list) - 1],
        marks={i: {"label": weeks_list[i], "style": {"writingMode": "vertical-rl"}} for i in range(0, len(weeks_list), max(1, len(weeks_list)//10))},
        tooltip={"placement": "bottom", "always_visible": True}
    )
], style={"width": f"{FILTER_CONFIG['week']['width_pct']}%", "display": "inline-block", "padding": "10px"}))

# Cause filter
if "cause" in FILTER_CONFIG:
    causes = df[COLUMNS["cause"]].dropna().unique() if LAYOUT["exclude_nan"] else df[COLUMNS["cause"]].unique()
    filter_elements.append(html.Div([
        html.Label("Причина ДТП:", style={"fontWeight": "bold"}),
        dcc.Dropdown(
            id="cause-filter",
            options=[{"label": c, "value": c} for c in sorted(causes)],
            placeholder="Выберите причину...",
            style={"width": "100%"}
        )
    ], style={"width": f"{FILTER_CONFIG['cause']['width_pct']}%", "display": "inline-block", "padding": "10px", "verticalAlign": "top"}))

# District filter
if "district" in FILTER_CONFIG:
    districts = df[COLUMNS["district"]].dropna().unique() if LAYOUT["exclude_nan"] else df[COLUMNS["district"]].unique()
    filter_elements.append(html.Div([
        html.Label("Район:", style={"fontWeight": "bold"}),
        dcc.Dropdown(
            id="district-filter",
            options=[{"label": d, "value": d} for d in sorted(districts)],
            placeholder="Выберите район...",
            style={"width": "100%"}
        )
    ], style={"width": "20%", "display": "inline-block", "padding": "10px", "verticalAlign": "top"}))

# Сортировка дашбордов: сначала "Top...", но "Top Branches" — последним среди Top
top_dashboards = [d for d in DASHBOARDS if d["title"].startswith("Top")]
other_dashboards = [d for d in DASHBOARDS if not d["title"].startswith("Top")]

# Переместить "Top Branches" в конец списка top_dashboards
top_dashboards_sorted = []
branches = None
for d in top_dashboards:
    if d["title"] == "Top Branches":
        branches = d
    else:
        top_dashboards_sorted.append(d)
if branches:
    top_dashboards_sorted.append(branches)

ordered_dashboards = top_dashboards_sorted + other_dashboards

# Дашборды (заглушки — будут обновляться через callback)
dashboard_placeholders = []
for i, dash in enumerate(ordered_dashboards):
    width = dash.get("width_pct") or LAYOUT["default_width_pct"]
    dashboard_placeholders.append(
        html.Div(id=f"chart-{i}", style={"width": f"{width}%", "display": "inline-block", "padding": "10px"})
    )

app.layout = html.Div([
    html.H2("Интерактивный дашборд", style={"textAlign": "center", "marginBottom": "20px"}),

    # Фильтры
    html.Div(filter_elements, style={"textAlign": "left", "marginBottom": "20px"}),

    # Графики
    html.Div(id="interactive-charts", children=dashboard_placeholders)
])

# =============================
# 6. CALLBACKS
# =============================

@app.callback(
    [Output(f"chart-{i}", "children") for i in range(len(ordered_dashboards))],
    Input("week-range", "value"),
    Input("cause-filter", "value"),
    Input("district-filter", "value")
)
def update_charts(week_range, cause_val, district_val):
    # Фильтрация по неделям
    from_idx, to_idx = week_range
    selected_weeks = weeks_list[from_idx:to_idx+1]
    filtered_df = df[df["WeekLabel"].isin(selected_weeks)].copy()

    # Фильтрация по другим полям
    if cause_val:
        filtered_df = filtered_df[filtered_df[COLUMNS["cause"]] == cause_val]
    if district_val:
        filtered_df = filtered_df[filtered_df[COLUMNS["district"]] == district_val]

    # Исключить NaN при необходимости
    if LAYOUT["exclude_nan"]:
        pass  # уже фильтруется в aggregate_top

    outputs = []
    for dash in ordered_dashboards:
        col = dash["data_column"]
        top_n = dash["top_n"]
        color = dash["color"]
        title = dash["title"]

        labels, values = aggregate_top(filtered_df, col, top_n=top_n, exclude_nan=LAYOUT["exclude_nan"])
        fig = create_bar_chart(labels, values, title, color)
        outputs.append(dcc.Graph(figure=fig))

    return outputs

# =============================
# 7. ЗАПУСК
# =============================

if __name__ == "__main__":
    app.run_server(debug=True)