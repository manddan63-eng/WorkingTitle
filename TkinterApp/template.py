import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd
import numpy as np
from datetime import datetime, date, time as dtTime, timedelta
import re
import openpyxl as ox
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import configparser
import json
from openpyxl.worksheet.table import Table, TableStyleInfo

 
def generate_weekly_dashboard_2025(self, journal_path):
        try:
            self.log("Генерация дашборда за 2025–2026 годы...")
            df = pd.read_excel(journal_path, sheet_name='Лист1')
            df['Дата ДТП'] = pd.to_datetime(df['Дата ДТП'], format='%d.%m.%Y', errors='coerce')
            df = df.dropna(subset=['Дата ДТП'])

            # === НАДЁЖНАЯ ЗАМЕНА NaN НА "Неизвестно" ===
            for col in ['Район', 'Округ', 'Причина ДТП', 'Название филиала']:
                df[col] = df[col].apply(lambda x: 'Неизвестно' if pd.isna(x) else str(x).strip())

            # === Подготовка данных для основных графиков (по годам и неделям) ===
            for year in [2025, 2026]:
                df_year = df[df['Дата ДТП'].dt.year == year].copy()
                df_year['Неделя'] = df_year['Дата ДТП'].dt.isocalendar().week

                total_all = df_year.groupby('Неделя').size()
                carrier_all = df_year[df_year['Виновник ДТП'] == 'Перевозчик'].groupby('Неделя').size()
                df1 = pd.concat([total_all, carrier_all], axis=1).fillna(0).astype(int)
                df1.columns = ['Всего ДТП', 'По вине перевозчика']
                df1 = df1.reset_index()

                df_with_victims = df_year[df_year['Кол-во пострадавших'] != 0]
                total_victims = df_with_victims.groupby('Неделя').size()
                carrier_victims = df_with_victims[df_with_victims['Виновник ДТП'] == 'Перевозчик'].groupby('Неделя').size()
                df2 = pd.concat([total_victims, carrier_victims], axis=1).fillna(0).astype(int)
                df2.columns = ['Всего ДТП', 'По вине перевозчика']
                df2 = df2.reset_index()

                df_without_victims = df_year[df_year['Кол-во пострадавших'] == 0]
                total_no_victims = df_without_victims.groupby('Неделя').size()
                carrier_no_victims = df_without_victims[df_without_victims['Виновник ДТП'] == 'Перевозчик'].groupby('Неделя').size()
                df3 = pd.concat([total_no_victims, carrier_no_victims], axis=1).fillna(0).astype(int)
                df3.columns = ['Всего ДТП', 'По вине перевозчика']
                df3 = df3.reset_index()

                prefix = f"_{year}"
                setattr(self, f'weeks1{prefix}', df1['Неделя'].tolist())
                setattr(self, f'total1{prefix}', df1['Всего ДТП'].tolist())
                setattr(self, f'rest1{prefix}', (df1['Всего ДТП'] - df1['По вине перевозчика']).tolist())
                setattr(self, f'carrier1{prefix}', df1['По вине перевозчика'].tolist())

                setattr(self, f'weeks2{prefix}', df2['Неделя'].tolist())
                setattr(self, f'total2{prefix}', df2['Всего ДТП'].tolist())
                setattr(self, f'rest2{prefix}', (df2['Всего ДТП'] - df2['По вине перевозчика']).tolist())
                setattr(self, f'carrier2{prefix}', df2['По вине перевозчика'].tolist())

                setattr(self, f'weeks3{prefix}', df3['Неделя'].tolist())
                setattr(self, f'total3{prefix}', df3['Всего ДТП'].tolist())
                setattr(self, f'rest3{prefix}', (df3['Всего ДТП'] - df3['По вине перевозчика']).tolist())
                setattr(self, f'carrier3{prefix}', df3['По вине перевозчика'].tolist())

            # === Топ филиалов (для статики и фильтрации) ===
            branch_mapping = {
                'Ф': 'Северный', 'ФСВ': 'Северо-Восточный', 'ФСЗ': 'Северо-Западный',
                'ФЮ': 'Южный', 'ФЮВ': 'Юго-Восточный', 'ФЮЗ': 'Юго-Западный',
                'ФВ': 'Восточный', 'ФЗ': 'Западный', 'ФЦ': 'Центральный',
                'ФС(Э)': 'Северный', 'ФСВ(Э)': 'Северо-Восточный', 'ФСЗ(Э)': 'Северо-Западный',
                'ФЮ(Э)': 'Южный', 'ФЮВ(Э)': 'Юго-Восточный', 'ФЮЗ(Э)': 'Юго-Западный',
                'ФВ(Э)': 'Восточный', 'ФЗ(Э)': 'Западный', 'ФЦ(Э)': 'Центральный',
                'ФС (Э)': 'Северный', 'ФСВ (Э)': 'Северо-Восточный', 'ФСЗ (Э)': 'Северо-Западный',
                'ФЮ (Э)': 'Южный', 'ФЮВ (Э)': 'Юго-Восточный', 'ФЮЗ (Э)': 'Юго-Западный',
                'ФВ (Э)': 'Восточный', 'ФЗ (Э)': 'Западный', 'ФЦ (Э)': 'Центральный',
                'ФС(э)': 'Северный', 'ФСВ(э)': 'Северо-Восточный', 'ФСЗ(э)': 'Северо-Западный',
                'ФЮ(э)': 'Южный', 'ФЮВ(э)': 'Юго-Восточный', 'ФЮЗ(э)': 'Юго-Западный',
                'ФВ(э)': 'Восточный', 'ФЗ(э)': 'Западный', 'ФЦ(э)': 'Центральный',
                'ФС (э)': 'Северный', 'ФСВ (э)': 'Северо-Восточный', 'ФСЗ (э)': 'Северо-Западный',
                'ФЮ (э)': 'Южный', 'ФЮВ (э)': 'Юго-Восточный', 'ФЮЗ (э)': 'Юго-Западный',
                'ФВ (э)': 'Восточный', 'ФЗ (э)': 'Западный', 'ФЦ (э)': 'Центральный',
                'ФС (Э)': 'Северный', 'ФСВ (Э)': 'Северо-Восточный', 'ФСЗ (Э)': 'Северо-Западный',
                'ФЮ (Э)': 'Южный', 'ФЮВ (Э)': 'Юго-Восточный', 'ФЮЗ (Э)': 'Юго-Западный',
                'ФВ (Э)': 'Восточный', 'ФЗ (Э)': 'Западный', 'ФЦ (Э)': 'Центральный',
            }
            df['Название филиала'] = df['Название филиала'].map(branch_mapping).fillna('Неизвестно')
            df['Название филиала'] = df['Название филиала'].apply(lambda x: 'Неизвестно' if pd.isna(x) else str(x).strip())

            # === Подготовка данных для интерактивных фильтров ===
            df_serial = df.copy()
            df_serial['Дата ДТП'] = df_serial['Дата ДТП'].dt.strftime('%Y-%m-%d')
            df_serial['Год'] = df['Дата ДТП'].dt.year
            df_serial['Неделя'] = df['Дата ДТП'].dt.isocalendar().week
            df_serial['Неделя_год'] = df_serial['Неделя'].astype(str) + ' нед ' + df_serial['Год'].astype(str) + ' года'

            weeks_list = (
                df_serial[['Год', 'Неделя', 'Неделя_год']]
                .drop_duplicates()
                .sort_values(['Год', 'Неделя'])['Неделя_год']
                .tolist()
            )

            districts = sorted(df['Район'].unique())
            causes = sorted(df['Причина ДТП'].unique())
            branches = sorted(df['Название филиала'].unique())

            data_json = json.dumps({
                'records': df_serial.to_dict(orient='records'),
                'filters': {
                    'districts': districts,
                    'causes': causes,
                    'branches': branches,
                    'weeks': weeks_list
                }
            }, ensure_ascii=False, default=str)
            all_weeks_1 = getattr(self, 'weeks1_2026', [])
            max_week_overall = max([4] + all_weeks_1)          
            html_content = f'''
    <!DOCTYPE html>
    <html lang="ru">
    <head>
    <meta charset="UTF-8">
    <title>ДТП по неделям — 2025/2026</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
    body {{ font-family: Arial, sans-serif; padding: 20px; background: #f9f9f9; }}
    .container {{ max-width: 1200px; margin: 0 auto; }}

    h1, h2 {{ text-align: center; color: #2c3e50; }}
    .overallStyle {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 30px; }}
    .filters {{ display: flex; gap: 15px; flex-wrap: wrap; align-items: end; justify-content: center; margin: 15px 0; }}
    .filters > div {{ display: flex; flex-direction: column; min-width: 150px; }}
    .filters label {{ margin-bottom: 4px; font-weight: bold; }}
    .filters select, .filters input[type="date"] {{ padding: 6px; border: 1px solid #ccc; border-radius: 4px; width: 100%; }}
    #cause-filter {{ width: 30vw !important; max-width: 300px; }}
    .filters button {{ align-self: flex-end; padding: 8px 16px; background: #2E86AB; color: white; border: none; border-radius: 4px; cursor: pointer; }}

    #interactive-charts {{
      display: flex;
      flex-wrap: wrap;
      gap: 20px;
      justify-content: center;
      width: 100%;
    }}

    .interactive-chart {{
      flex: 1 1 calc(50% - 20px);
      min-width: 200px;
      height: 190px;
      background: white;
      padding: 10px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      box-sizing: border-box;
    }}
    </style>
    </head>
    <body>

    <!-- === ФИЛЬТРЫ В САМОМ НАЧАЛЕ === -->
    <div class="container">
      <h2>Фильтры по данным</h2>
      <div class="filters">
        <div><label>Неделя с:</label><select id="week-from"><option value="">Выберите</option>{''.join(f'<option value="{i}">{w}</option>' for i, w in enumerate(weeks_list))}</select></div>
        <div><label>Неделя по:</label><select id="week-to"><option value="">Выберите</option>{''.join(f'<option value="{i}">{w}</option>' for i, w in enumerate(weeks_list))}</select></div>
        <div><label>Район:</label><select id="district-filter"><option value="">Все районы</option>{''.join(f'<option value="{d}">{d}</option>' for d in districts)}</select></div>
        <div><label>Причина ДТП:</label><select id="cause-filter"><option value="">Все причины</option>{''.join(f'<option value="{c}">{c}</option>' for c in causes)}</select></div>
        <button onclick="updateAllCharts()">Применить</button>
      </div>
    </div>

    <!-- === ИНТЕРАКТИВНЫЕ ГРАФИКИ (4 штуки) === -->
    <div class="container">
      <h2>Обзор по категориям (с фильтрацией)</h2>
      <div id="interactive-charts"></div>
    </div>

    <!-- === СТАТИЧНЫЕ ГРАФИКИ ПО НЕДЕЛЯМ === -->
    <div class="container">
      <h1>ДТП по неделям — 2025/2026 год</h1>
      <div id="chart" class="overallStyle"></div>
    </div>
    <div class="container">
      <h1>ДТП по вине перевозчика — с пострадавшими</h1>
      <div id="chart2" class="overallStyle"></div>
    </div>
    <div class="container">
      <h1>ДТП по вине перевозчика — без пострадавших</h1>
      <div id="chart3" class="overallStyle"></div>
    </div>

    <script>
    // === СТАТИЧНЫЕ ДАННЫЕ ===

    const maxWeekOverall = {json.dumps(max_week_overall)};
    const weeks1_2025 = {json.dumps(getattr(self, 'weeks1_2025'))};
    const total1_2025 = {json.dumps(getattr(self, 'total1_2025'))};
    const carrier1_2025 = {json.dumps(getattr(self, 'carrier1_2025'))};
    const rest1_2025 = {json.dumps(getattr(self, 'rest1_2025'))};

    const weeks2_2025 = {json.dumps(getattr(self, 'weeks2_2025'))};
    const total2_2025 = {json.dumps(getattr(self, 'total2_2025'))};
    const carrier2_2025 = {json.dumps(getattr(self, 'carrier2_2025'))};
    const rest2_2025 = {json.dumps(getattr(self, 'rest2_2025'))};

    const weeks3_2025 = {json.dumps(getattr(self, 'weeks3_2025'))};
    const total3_2025 = {json.dumps(getattr(self, 'total3_2025'))};
    const carrier3_2025 = {json.dumps(getattr(self, 'carrier3_2025'))};
    const rest3_2025 = {json.dumps(getattr(self, 'rest3_2025'))};

    const weeks1_2026 = {json.dumps(getattr(self, 'weeks1_2026'))};
    const total1_2026 = {json.dumps(getattr(self, 'total1_2026'))};
    const carrier1_2026 = {json.dumps(getattr(self, 'carrier1_2026'))};
    const rest1_2026 = {json.dumps(getattr(self, 'rest1_2026'))};

    const weeks2_2026 = {json.dumps(getattr(self, 'weeks2_2026'))};
    const total2_2026 = {json.dumps(getattr(self, 'total2_2026'))};
    const carrier2_2026 = {json.dumps(getattr(self, 'carrier2_2026'))};
    const rest2_2026 = {json.dumps(getattr(self, 'rest2_2026'))};

    const weeks3_2026 = {json.dumps(getattr(self, 'weeks3_2026'))};
    const total3_2026 = {json.dumps(getattr(self, 'total3_2026'))};
    const carrier3_2026 = {json.dumps(getattr(self, 'carrier3_2026'))};
    const rest3_2026 = {json.dumps(getattr(self, 'rest3_2026'))};

    const weeksList = {json.dumps(weeks_list)};
    const rawData = {data_json};

    // === АГРЕГАЦИЯ ===
    function aggregate(data, key, topN = null) {{
      const counts = {{}};
      data.forEach(row => {{
        const val = row[key];
        if (val !== null && val !== undefined && val !== "") {{
          counts[val] = (counts[val] || 0) + 1;
        }}
      }});
      let entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      if (topN) entries = entries.slice(0, topN);
      return entries;
    }}

    // === ФИЛЬТРАЦИЯ ПО ДИАПАЗОНУ НЕДЕЛЬ ===
    function filterData() {{
      const fromIdx = document.getElementById('week-from').value;
      const toIdx = document.getElementById('week-to').value;
      const district = document.getElementById('district-filter').value;
      const cause = document.getElementById('cause-filter').value;

      let filtered = rawData.records;

      if (fromIdx !== '' || toIdx !== '') {{
        const fromWeek = fromIdx !== '' ? weeksList[parseInt(fromIdx)] : null;
        const toWeek = toIdx !== '' ? weeksList[parseInt(toIdx)] : null;
        filtered = filtered.filter(row => {{
          const weekYear = row['Неделя'].toString() + ' нед ' + row['Год'].toString() + ' года';
          const idx = weeksList.indexOf(weekYear);
          if (idx === -1) return false;
          if (fromIdx !== '' && idx < parseInt(fromIdx)) return false;
          if (toIdx !== '' && idx > parseInt(toIdx)) return false;
          return true;
        }});
      }}

      filtered = filtered.filter(row =>
        (!district || row['Район'] === district) &&
        (!cause || row['Причина ДТП'] === cause)
      );

      return filtered;
    }}

    // === ОБНОВЛЕНИЕ ВСЕХ ИНТЕРАКТИВНЫХ ГРАФИКОВ ===
    function updateAllCharts() {{
      const container = document.getElementById('interactive-charts');
      container.innerHTML = '';
      const currentData = filterData();

      // Топ причин
      const causes = aggregate(currentData, 'Причина ДТП', 10);
      if (causes.length > 0) {{
        const div = document.createElement('div');
        div.className = 'interactive-chart';
        Plotly.newPlot(div, [{{
          x: causes.map(d => d[0]),
          y: causes.map(d => d[1]),
          type: 'bar',
          marker: {{ color: '#C73E1D' }}
        }}], {{
          title: 'Топ причин ДТП',
          margin: {{ t: 25, b: 60, l: 45, r: 15 }},
          xaxis: {{ tickangle: -45 }},
          height: 180,
          width: 510
        }});
        container.appendChild(div);
      }}

      // Топ районов
      const districts = aggregate(currentData, 'Район', 10);
      if (districts.length > 0) {{
        const div = document.createElement('div');
        div.className = 'interactive-chart';
        Plotly.newPlot(div, [{{
          x: districts.map(d => d[0]),
          y: districts.map(d => d[1]),
          type: 'bar',
          marker: {{ color: '#3BB273' }}
        }}], {{
          title: 'Топ районов',
          margin: {{ t: 25, b: 60, l: 45, r: 15 }},
          xaxis: {{ tickangle: -45 }},
          height: 180,
          width: 510
        }});
        container.appendChild(div);
      }}

      // Округа
      const okrugs = aggregate(currentData, 'Округ');
      if (okrugs.length > 0) {{
        const div = document.createElement('div');
        div.className = 'interactive-chart';
        Plotly.newPlot(div, [{{
          x: okrugs.map(d => d[0]),
          y: okrugs.map(d => d[1]),
          type: 'bar',
          marker: {{ color: '#5D5D81' }}
        }}], {{
          title: 'ДТП по округам',
          margin: {{ t: 25, b: 60, l: 45, r: 15 }},
          xaxis: {{ tickangle: -45 }},
          height: 180,
          width: 510
        }});
        container.appendChild(div);
      }}

      // Топ филиалов (с фильтрацией!)
      const branches = aggregate(currentData, 'Название филиала', 10);
      if (branches.length > 0) {{
        const div = document.createElement('div');
        div.className = 'interactive-chart';
        Plotly.newPlot(div, [{{
          x: branches.map(d => d[0]),
          y: branches.map(d => d[1]),
          type: 'bar',
          marker: {{ color: '#8B4513' }}
        }}], {{
          title: 'Топ филиалов (с фильтрацией)',
          margin: {{ t: 25, b: 60, l: 45, r: 15 }},
          xaxis: {{ tickangle: -45 }},
          height: 180,
          width: 510
        }});
        container.appendChild(div);
      }}
    }}

    // === СТАТИЧНЫЕ ГРАФИКИ ПО НЕДЕЛЯМ ===
    // График 1
    const x_2025_1 = weeks1_2025.map(w => w - 0.15);
    const x_2026_1 = weeks1_2026.map(w => w + 0.15);
    Plotly.newPlot('chart', [
      {{ x: x_2025_1, y: carrier1_2025, type: 'bar', name: '2025: По вине перевозчика', marker: {{ color: '#C73E1D' }}, textposition: 'inside', textfont: {{ color: 'white' }},hovertemplate: '2025: По вине перевозчика: %{{y}}<extra></extra>' }},
      {{ x: x_2025_1, y: rest1_2025, type: 'bar',  name: '2025: Всего ДТП', marker: {{ color: '#2E86AB' }}, text: total1_2025.map(val => '2025: ' + val),customdata: total1_2025,hovertemplate: '2025: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }} }},
      {{ x: x_2026_1, y: carrier1_2026, type: 'bar', name: '2026: По вине перевозчика', marker: {{ color: '#A52A2A' }}, textposition: 'inside', textfont: {{ color: 'white' }} ,hovertemplate: '2026: По вине перевозчика: %{{y}}<extra></extra>'}},
      {{ x: x_2026_1, y: rest1_2026, type: 'bar',  name: '2026: Всего ДТП', marker: {{ color: '#1E5E8C' }}, text: total1_2026.map(val => '2026: ' + val),customdata: total1_2026,hovertemplate: '2026: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }}  }}
    ], {{
      title: 'ДТП по неделям',
      xaxis: {{ title: 'Неделя года', tickmode: 'array', tickvals: Array.from({{length: 53}}, (_, i) => i + 1), range: [maxWeekOverall - 4 + 0.5, maxWeekOverall + 0.5] }},
      yaxis: {{ title: 'Количество ДТП' }},
      barmode: 'stack',
      hovermode: 'x unified',
      legend: {{ orientation: 'h', yanchor: 'bottom', y: 1.02, xanchor: 'right', x: 1 }}
    }});

    // График 2
    const x_2025_2 = weeks2_2025.map(w => w - 0.15);
    const x_2026_2 = weeks2_2026.map(w => w + 0.15);
    Plotly.newPlot('chart2', [
      {{ x: x_2025_2, y: carrier2_2025, type: 'bar', name: '2025: По вине перевозчика', marker: {{ color: '#C73E1D' }},  textposition: 'inside', textfont: {{ color: 'white' }},hovertemplate: '2025: По вине перевозчика: %{{y}}<extra></extra>' }},
      {{ x: x_2025_2, y: rest2_2025, type: 'bar',  name: '2025: Всего ДТП', marker: {{ color: '#2E86AB' }} , text: total2_2025.map(val => '2025: ' + val),customdata: total2_2025,hovertemplate: '2025: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }}}},
      {{ x: x_2026_2, y: carrier2_2026, type: 'bar', name: '2026: По вине перевозчика', marker: {{ color: '#A52A2A' }},  textposition: 'inside', textfont: {{ color: 'white' }},hovertemplate: '2026: По вине перевозчика: %{{y}}<extra></extra>' }},
      {{ x: x_2026_2, y: rest2_2026, type: 'bar',  name: '2026: Всего ДТП', marker: {{ color: '#1E5E8C' }} , text: total2_2026.map(val => '2026: ' + val),customdata: total2_2026,hovertemplate: '2026: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }}}}
    ], {{
      title: 'ДТП по вине перевозчика — с пострадавшими',
      xaxis: {{ title: 'Неделя года', tickmode: 'array', tickvals: Array.from({{length: 53}}, (_, i) => i + 1), range: [maxWeekOverall - 4 + 0.5, maxWeekOverall + 0.5] }},
      yaxis: {{ title: 'Количество ДТП' }},
      barmode: 'stack',
      hovermode: 'x unified'
    }});

    // График 3
    const x_2025_3 = weeks3_2025.map(w => w - 0.15);
    const x_2026_3 = weeks3_2026.map(w => w + 0.15);
    Plotly.newPlot('chart3', [
      {{ x: x_2025_3, y: carrier3_2025, type: 'bar', name: '2025: По вине перевозчика', marker: {{ color: '#C73E1D' }}, textposition: 'inside', textfont: {{ color: 'white' }} ,hovertemplate: '2025: По вине перевозчика: %{{y}}<extra></extra>'}},
      {{ x: x_2025_3, y: rest3_2025, type: 'bar',  name: '2025: Всего ДТП', marker: {{ color: '#2E86AB' }},text: total3_2025.map(val => '2025: ' + val),customdata: total3_2025,hovertemplate: '2025: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }} }},
      {{ x: x_2026_3, y: carrier3_2026, type: 'bar', name: '2026: По вине перевозчика', marker: {{ color: '#A52A2A' }},  textposition: 'inside', textfont: {{ color: 'white' }} ,hovertemplate: '2026: По вине перевозчика: %{{y}}<extra></extra>'}},
      {{ x: x_2026_3, y: rest3_2026, type: 'bar',  name: '2026: Всего ДТП', marker: {{ color: '#1E5E8C' }} , text: total3_2026.map(val => '2026: ' + val),customdata: total3_2026,hovertemplate: '2026: Всего ДТП: %{{customdata}}<extra></extra>', textposition: 'outside', textfont: {{ color: 'black' }}}}
    ], {{
      title: 'ДТП по вине перевозчика — без пострадавших',
      xaxis: {{ title: 'Неделя года', tickmode: 'array', tickvals: Array.from({{length: 53}}, (_, i) => i + 1), range: [maxWeekOverall - 4 + 0.5, maxWeekOverall + 0.5] }},
      yaxis: {{ title: 'Количество ДТП' }},
      barmode: 'stack',
      hovermode: 'x unified'
    }});

    // Первичная отрисовка
    updateAllCharts();
    </script>
    </body>
    </html>
    '''
            output_html = os.path.join(os.path.dirname(journal_path), 'Dashboard_2025_2026_weekly.html')
            with open(output_html, 'w', encoding='utf-8') as f:
                f.write(html_content)
            self.log(f"Дашборд сохранён: {os.path.basename(output_html)}")

        except Exception as e:
            self.log(f"Ошибка при создании дашборда: {e}")''


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()