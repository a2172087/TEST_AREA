import sys
import os
import pandas as pd
import re
import numpy as np
import math
from datetime import datetime
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt, QUrl, QTimer, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QMessageBox, QProgressBar, QLabel,
                           QHBoxLayout, QFileDialog, QComboBox)
from PyQt5.QtWebEngineWidgets import QWebEngineView
import tempfile
import webbrowser
import subprocess
import pyodbc
import json
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import J750_J750EX_UFLEX_process_V3
import ETS88_Accotest_process_V3
import AG93000_process_V4
import T2K_process_V1
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# 獲取使用者名稱
username = os.environ.get('USERNAME', 'Unknown')

# 從JSON檔案讀取SQL Server連線資訊
json_path = r"M:\BI_Database\Apps\Database\Apps_Database\O_All\SQL_Server\SQL_Server_Info_User_BI.json"
with open(json_path, 'r') as file:
    sql_connection_info = json.load(file)

# SQL Server連線資訊
SQL_SERVER_INFO = {
    "server": sql_connection_info["server"],
    "database": sql_connection_info["database"], 
    "username": sql_connection_info["username"],
    "password": sql_connection_info["password"],
    "apps_log_table": sql_connection_info["apps_log_table"]  
}

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

STYLE_SHEET = """
QWidget {
    background-color: #2D2D2D;
    color: #E0E0E0;
    font-family: "微軟正黑體";
    font-size: 9pt;
}

QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #454545, stop:1 #383838);
    border: 1px solid #454545;
    border-radius: 6px;
    padding: 8px 16px;
    color: #E0E0E0;
    min-width: 120px;
    min-height: 35px;
    margin: 2px;
}

QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #505050, stop:1 #404040);
    border: 1px solid #666666;
    color: #FFFFFF;
}

QPushButton:pressed {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #353535, stop:1 #313131);
    border: 1px solid #444444;
}

QPushButton:disabled {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #2D2D2D, stop:1 #282828);
    border: 1px solid #333333;
    color: #666666;
}

QProgressBar {
    border: none;
    border-radius: 4px;
    text-align: center;
    padding: 1px;
    background-color: #333333;
    height: 12px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #505050, stop:1 #404040);
    border-radius: 3px;
}

QProgressBar:hover {
    background-color: #383838;
}

QLabel {
    color: #E0E0E0;
    padding: 2px;
}

QComboBox {
    background-color: #383838;
    border: 1px solid #454545;
    border-radius: 4px;
    color: #E0E0E0;
    padding: 4px 8px;
    min-height: 25px;
}

QComboBox:hover {
    background-color: #404040;
    border: 1px solid #666666;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left: none;  /* 移除分隔線 */
    border-top-right-radius: 4px;
    border-bottom-right-radius: 4px;
}

QComboBox QAbstractItemView {
    background-color: #2D2D2D;
    border: 1px solid #454545;
    selection-background-color: #505050;
    selection-color: #FFFFFF;
}
"""

# 報告生成相關函數
def create_line_chart(wafer_data, axis_type, standard_value=None, standard_point_data=None):
    """為指定的軸類型創建折線圖"""
    
    fig = go.Figure()
    
    # 跟踪最大/最小值用於註釋定位
    max_y_value = float('-inf')
    min_y_value = float('inf')
    
    # 為 x 軸創建連續索引
    continuous_x = []
    continuous_y = []
    wafer_boundaries = []
    wafer_labels = []
    
    # 為每個數據點創建標籤列表
    point_labels = []
    
    # 根據軸類型獲取值
    value_key = f"{axis_type}_values"
    
    # 按開始時間排序晶圓
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    
    current_index = 0  #從索引 0 開始，不再為 AutoZ complete 點預留索引
    
    # 處理每個晶圓
    for wafer_id, data in sorted_wafers:
        if value_key in data and data[value_key]:
            values = data[value_key]
            
            # 此晶圓的起始和結束索引
            start_idx = current_index
            end_idx = start_idx + len(values)
            
            # 添加到連續數組
            continuous_x.extend(range(start_idx, end_idx))
            continuous_y.extend(values)
            
            # 為每個數據點分配晶圓ID作為標籤
            point_labels.extend([wafer_id] * len(values))
            
            # 存儲晶圓邊界信息
            wafer_boundaries.append((start_idx, end_idx - 1))
            wafer_labels.append(wafer_id)
            
            # 更新當前索引
            current_index = end_idx
            
            # 更新最大/最小值
            if values:
                max_y_value = max(max_y_value, max(values))
                min_y_value = min(min_y_value, min(values))
    
    # 不再單獨插入 AutoZ complete 點，而是識別第一個點是否為 AutoZ complete
    # 準備顏色和大小數組
    if continuous_x and continuous_y:
        # 為每個點設定顏色
        colors = []
        sizes = []
        
        # 檢查第一個點是否為 Auto Z complete 點
        is_first_point_autoz = False
        if standard_point_data and axis_type in standard_point_data:
            autoz_value = standard_point_data[axis_type]
            # 通過比對座標值判斷（容許 0.001 的誤差）
            if len(continuous_y) > 0 and abs(continuous_y[0] - autoz_value) < 0.001:
                is_first_point_autoz = True
                # 更新最大/最小值以包含 AutoZ complete 點
                max_y_value = max(max_y_value, autoz_value)
                min_y_value = min(min_y_value, autoz_value)
        
        for i in range(len(continuous_x)):
            if i == 0 and is_first_point_autoz:
                # 第一個點是 AutoZ complete 點，用紅色和較大尺寸
                colors.append('#e5857b')
                sizes.append(20)
                # 修改標籤
                point_labels[0] = "AutoZ Complete"
            else:
                # 其他點用原來的顏色和尺寸
                colors.append('#93A1C1')
                sizes.append(8)
        
        # 添加主線跡與標記
        fig.add_trace(
            go.Scatter(
                x=continuous_x,
                y=continuous_y,
                mode='lines+markers',
                name=f"{axis_type.upper()} Values",
                text=point_labels,
                line=dict(color='#93A1C1', width=4),
                marker=dict(
                    size=sizes,
                    color=colors,
                    line=dict(color='white', width=2)
                ),
                hovertemplate="Point: %{text}<br>" + axis_type.upper() + " Value: %{y:.2f} µm<extra></extra>"
            )
        )
    
    # 僅為 Z 軸添加標準參考線
    if standard_value is not None and axis_type == 'z':
        x_range_start = min(continuous_x) if continuous_x else 0
        x_range_end = max(continuous_x) if continuous_x else 1
        
        fig.add_trace(
            go.Scatter(
                x=[x_range_start, x_range_end],
                y=[standard_value, standard_value],
                mode='lines',
                name=f"{axis_type.upper()} Standard",
                line=dict(color='#e5857b', width=4, dash='dash')
            )
        )
        
        # 在 Z 標準線上添加常量標籤
        if continuous_x:
            # 在圖表的中間位置添加標籤
            x_range = x_range_end - x_range_start if x_range_end > x_range_start else 1
            x_pos = x_range_start + x_range / 2
            
            fig.add_annotation(
                x=x_pos,
                y=standard_value,
                text=f"Z Standard: {standard_value:.2f} µm",
                showarrow=False,
                yshift=15,
                bgcolor="rgba(255, 255, 255, 0.8)",
                bordercolor="#e5857b",
                borderwidth=2,
                borderpad=4,
                font=dict(color="#e5857b", size=12, family="Microsoft JhengHei")
            )
    
    # 添加晶圓邊界標記和註釋
    for i, (start, end) in enumerate(wafer_boundaries):
        # 在晶圓邊界添加半透明垂直線
        fig.add_trace(
            go.Scatter(
                x=[start, start],
                y=[min_y_value, max_y_value],
                mode='lines',
                line=dict(
                    color='rgba(69, 73, 106, 0.25)',
                    width=1.2,  
                    dash='dot'
                ),
                showlegend=False
            )
        )
    
    # 更新佈局
    fig.update_layout(
        width=1100,
        height=600,
        title=dict(
            text=f"{axis_type.upper()} AutoZ Values",
            x=0.5,
            y=0.98,
            xanchor='center',
            yanchor='top',
            font=dict(family='Microsoft JhengHei', size=18, weight='bold')
        ),
        xaxis=dict(
            title="Sequential Index",
            title_font=dict(family='Microsoft JhengHei', size=14, weight='bold'),
            tickfont=dict(family='Microsoft JhengHei', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(
            title=f"{axis_type.upper()} Value (µm)",
            title_font=dict(family='Microsoft JhengHei', size=14, weight='bold'),
            tickfont=dict(family='Microsoft JhengHei', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        legend=dict(
            x=1.1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0.8)',
            bordercolor='lightgray',
            borderwidth=1,
            font=dict(family='Microsoft JhengHei', size=12)
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        hovermode='closest',
        margin=dict(l=50, r=30, t=60, b=50)
    )
    
    # 計算統計數據（包含 AutoZ complete 點）
    all_values = continuous_y.copy() if continuous_y else []
    
    stats = {
        'min': min(all_values) if all_values else 0,
        'max': max(all_values) if all_values else 0,
        'mean': np.mean(all_values) if all_values else 0,
        'median': np.median(all_values) if all_values else 0,
        'std': np.std(all_values) if all_values else 0,
        'count': len(all_values)
    }
    
    return fig, stats

def create_wafer_status_dashboard(wafer_data, z_standard):
    """Create a wafer status dashboard showing which wafers have Z values below standard
    
    Args:
        wafer_data: Dictionary containing wafer data
        z_standard: Z standard value
        
    Returns:
        HTML content for the wafer status dashboard
    """
    # 檢查每個晶圓的Z值是否有低於標準的點
    wafer_status = {}
    
    # 按開始時間排序晶圓
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    
    for wafer_id, data in sorted_wafers:
        if 'z_values' in data and data['z_values']:
            # 檢查是否有任何Z值低於標準
            below_standard = any(z < z_standard for z in data['z_values'])
            # 計算低於標準的點數
            below_count = sum(1 for z in data['z_values'] if z < z_standard)
            # 計算總點數
            total_count = len(data['z_values'])
            # 計算百分比
            percent_below = (below_count / total_count * 100) if total_count > 0 else 0
            
            # 存儲狀態信息
            wafer_status[wafer_id] = {
                'below_standard': below_standard,
                'below_count': below_count,
                'total_count': total_count,
                'percent_below': percent_below
            }
    
    # 計算每行顯示的晶圓數量（改為4個）
    wafers_per_row = 4
    
    # 創建HTML內容
    html_content = '''
    <div class="dashboard-container">
        <h2 class="dashboard-title">Wafer Status Dashboard</h2>
        <p class="dashboard-description">Status shows wafers with Z values below standard (red = has points below standard)</p>
        <div class="wafer-grid">
    '''
    
    # 計算行數
    row_count = math.ceil(len(wafer_status) / wafers_per_row)
    
    # 添加晶圓狀態卡片
    wafer_count = 0
    for wafer_id, status in wafer_status.items():
        # 確定卡片顏色
        card_class = "wafer-card-red" if status['below_standard'] else "wafer-card-green"
        
        # 如果是行首，添加新行開始標記
        if wafer_count % wafers_per_row == 0:
            html_content += '<div class="wafer-row">'
        
        # 添加晶圓卡片
        html_content += f'''
        <div class="wafer-card {card_class}">
            <div class="wafer-id">{wafer_id}</div>
            <div class="wafer-stats">
                <div class="stat-item">Below Standard: <span class="stat-value">{status['below_count']}/{status['total_count']}</span></div>
                <div class="stat-item">Percentage: <span class="stat-value">{status['percent_below']:.1f}%</span></div>
            </div>
        </div>
        '''
        
        wafer_count += 1
        
        # 如果是行尾或最後一個晶圓，添加行結束標記
        if wafer_count % wafers_per_row == 0 or wafer_count == len(wafer_status):
            html_content += '</div>'
    
    # 關閉容器
    html_content += '''
        </div>
    </div>
    '''
    
    # 添加CSS樣式
    css = '''
    <style>
        .dashboard-container {
            padding: 20px;
            background-color: #f8f9fa;
            border-radius: 8px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            max-width: 1200px;
            margin: 20px auto;
        }
        
        .dashboard-title {
            text-align: center;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 5px;
            color: #333;
            font-family: "Microsoft JhengHei", Arial, sans-serif;
        }
        
        .dashboard-description {
            text-align: center;
            font-size: 14px;
            color: #666;
            margin-bottom: 20px;
            font-family: "Microsoft JhengHei", Arial, sans-serif;
        }
        
        .wafer-grid {
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        
        .wafer-row {
            display: flex;
            justify-content: center;
            gap: 15px;
        }
        
        .wafer-card {
            padding: 15px;
            border-radius: 8px;
            width: 220px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .wafer-card-green {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
        }
        
        .wafer-card-red {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
        }
        
        .wafer-id {
            font-weight: bold;
            font-size: 18px;
            text-align: center;
            margin-bottom: 10px;
            font-family: "Microsoft JhengHei", Arial, sans-serif;
        }
        
        .wafer-stats {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }
        
        .stat-item {
            font-size: 14px;
            font-family: "Microsoft JhengHei", Arial, sans-serif;
            white-space: nowrap;
        }
        
        .stat-value {
            font-weight: bold;
        }
    </style>
    '''
    
    return css + html_content

def create_z_anomaly_chart(wafer_data, z_standard, standard_point_data=None):
    """Create a chart highlighting Z values below standard
    
    Args:
        wafer_data: Dictionary containing wafer data
        z_standard: Z standard value
        standard_point_data: AutoZ complete point data
        
    Returns:
        Plotly figure and statistics
    """
    fig = go.Figure()
    
    # 準備數據
    continuous_x = []
    normal_y = []
    anomaly_y = []
    normal_indices = []
    anomaly_indices = []
    wafer_ids = []
    
    # 按開始時間排序晶圓
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    
    current_index = 0  # 從索引 0 開始
    
    # 處理每個晶圓
    for wafer_id, data in sorted_wafers:
        if 'z_values' in data and data['z_values']:
            values = data['z_values']
            
            # 為這個晶圓創建索引
            indices = list(range(current_index, current_index + len(values)))
            
            # 將數據點分為正常和異常
            for idx, z_val in zip(indices, values):
                wafer_ids.append(wafer_id)
                if z_val < z_standard:
                    anomaly_indices.append(idx)
                    anomaly_y.append(z_val)
                else:
                    normal_indices.append(idx)
                    normal_y.append(z_val)
            
            # 更新索引
            current_index += len(values)
    
    # 不再單獨添加 AutoZ complete 點，而是識別第一個點
    # 檢查第一個點是否為 Auto Z complete 點
    is_first_point_autoz = False
    if standard_point_data and 'z' in standard_point_data:
        autoz_z_value = standard_point_data['z']
        if len(wafer_ids) > 0:
            # 檢查第一個點的值
            if normal_indices and normal_indices[0] == 0:
                first_z = normal_y[0]
            elif anomaly_indices and anomaly_indices[0] == 0:
                first_z = anomaly_y[0]
            else:
                first_z = None
            
            if first_z is not None and abs(first_z - autoz_z_value) < 0.001:
                is_first_point_autoz = True
    
    # 添加正常點
    if normal_indices:
        normal_wafer_ids_list = []
        for i in normal_indices:
            if i < len(wafer_ids):
                normal_wafer_ids_list.append(wafer_ids[i])
        
        # 為正常點準備顏色和尺寸
        normal_colors = []
        normal_sizes = []
        normal_symbols = []
        
        for i, idx in enumerate(normal_indices):
            if idx == 0 and is_first_point_autoz:
                # 第一個點是 AutoZ complete 點
                normal_colors.append('#FF6600')  # 橙色
                normal_sizes.append(14)
                normal_symbols.append('diamond')
            else:
                # 一般正常點
                normal_colors.append('#4CAF50')
                normal_sizes.append(8)
                normal_symbols.append('circle')
        
        fig.add_trace(
            go.Scatter(
                x=normal_indices,
                y=normal_y,
                mode='markers',
                name='Normal Points',
                marker=dict(
                    color=normal_colors,
                    size=normal_sizes,
                    symbol=normal_symbols,
                    line=dict(width=1, color='white')
                ),
                text=[f"{'AutoZ Complete' if (idx == 0 and is_first_point_autoz) else f'Wafer ID: {wid}'}<br>Z Value: {z:.2f} µm" 
                      for idx, wid, z in zip(normal_indices, normal_wafer_ids_list, normal_y)],
                hoverinfo='text'
            )
        )
    
    # 添加異常點
    if anomaly_indices:
        anomaly_wafer_ids_list = []
        for i in anomaly_indices:
            if i < len(wafer_ids):
                anomaly_wafer_ids_list.append(wafer_ids[i])
        
        # 為異常點準備顏色和尺寸
        anomaly_colors = []
        anomaly_sizes = []
        anomaly_symbols = []
        
        for i, idx in enumerate(anomaly_indices):
            if idx == 0 and is_first_point_autoz:
                # 第一個點是 AutoZ complete 點但低於標準
                anomaly_colors.append('#FF6600')  # 橙色
                anomaly_sizes.append(14)
                anomaly_symbols.append('diamond')
            else:
                # 一般異常點
                anomaly_colors.append('#F44336')
                anomaly_sizes.append(10)
                anomaly_symbols.append('circle')
        
        fig.add_trace(
            go.Scatter(
                x=anomaly_indices,
                y=anomaly_y,
                mode='markers',
                name='Below Standard',
                marker=dict(
                    color=anomaly_colors,
                    size=anomaly_sizes,
                    symbol=anomaly_symbols,
                    line=dict(width=1, color='white')
                ),
                text=[f"{'AutoZ Complete' if (idx == 0 and is_first_point_autoz) else f'Wafer ID: {wid}'}<br>Z Value: {z:.2f} µm" 
                      for idx, wid, z in zip(anomaly_indices, anomaly_wafer_ids_list, anomaly_y)],
                hoverinfo='text'
            )
        )
    
    # 添加標準線
    x_range_start = 0
    x_range_end = current_index
    fig.add_trace(
        go.Scatter(
            x=[x_range_start, x_range_end],
            y=[z_standard, z_standard],
            mode='lines',
            name=f"Z Standard ({z_standard} µm)",
            line=dict(color='#E91E63', width=3, dash='dash')
        )
    )
    
    # 計算異常點統計
    total_points = len(normal_indices) + len(anomaly_indices)
    anomaly_count = len(anomaly_indices)
    anomaly_percent = (anomaly_count / total_points * 100) if total_points > 0 else 0
    
    # 更新佈局
    fig.update_layout(
        width=1100,
        height=600,
        title=dict(
            text=f"Z Value Anomaly Analysis (Below Standard: {anomaly_count}/{total_points}, {anomaly_percent:.1f}%)",
            x=0.5,
            y=0.98,
            xanchor='center',
            yanchor='top',
            font=dict(family='Microsoft JhengHei', size=18, weight='bold')
        ),
        xaxis=dict(
            title="Sequential Index",
            title_font=dict(family='Microsoft JhengHei', size=14, weight='bold'),
            tickfont=dict(family='Microsoft JhengHei', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(
            title="Z Value (µm)",
            title_font=dict(family='Microsoft JhengHei', size=14, weight='bold'),
            tickfont=dict(family='Microsoft JhengHei', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        legend=dict(
            x=1.1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0.8)',
            bordercolor='lightgray',
            borderwidth=1,
            font=dict(family='Microsoft JhengHei', size=12)
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        hovermode='closest',
        margin=dict(l=50, r=30, t=60, b=50)
    )
    
    # 統計數據
    stats = {
        'total_points': total_points,
        'normal_points': len(normal_indices),
        'anomaly_points': anomaly_count,
        'anomaly_percent': anomaly_percent
    }
    
    return fig, stats

def generate_html_report(wafer_data, x_standard, y_standard, z_standard, machine_type):
    """生成包含所有視圖的 HTML 報告
    
    Args:
        wafer_data: 晶圓數據
        x_standard: X軸標準值
        y_standard: Y軸標準值
        z_standard: Z軸標準值
        machine_type: 機台類型 (用於標題)
    """
    # 準備AutoZ complete點位數據
    standard_point_data = {
        'x': x_standard,
        'y': y_standard,
        'z': z_standard
    }
    
    # 為所有三個軸生成圖表（包含AutoZ complete點位）
    x_fig, x_stats = create_line_chart(wafer_data, 'x', x_standard, standard_point_data)
    y_fig, y_stats = create_line_chart(wafer_data, 'y', y_standard, standard_point_data)
    z_fig, z_stats = create_line_chart(wafer_data, 'z', z_standard, standard_point_data)
    
    # 生成新的圖表（包含AutoZ complete點位）
    z_anomaly_fig, z_anomaly_stats = create_z_anomaly_chart(wafer_data, z_standard, standard_point_data)
    wafer_status_html = create_wafer_status_dashboard(wafer_data, z_standard)
    
    # 轉換為 HTML
    x_html = x_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    y_html = y_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    z_html = z_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    z_anomaly_html = z_anomaly_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 創建統計 HTML
    x_stats_html = f'''
    <div class="statistics-box">
        <div class="stat-item"><span class="stat-label">X Min:</span> <span class="stat-value">{x_stats['min']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">X Max:</span> <span class="stat-value">{x_stats['max']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">X Mean:</span> <span class="stat-value">{x_stats['mean']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">X Median:</span> <span class="stat-value">{x_stats['median']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">X Std Dev:</span> <span class="stat-value">{x_stats['std']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Data Points:</span> <span class="stat-value">{x_stats['count']:,}</span></div>
    </div>
    '''
    
    y_stats_html = f'''
    <div class="statistics-box">
        <div class="stat-item"><span class="stat-label">Y Min:</span> <span class="stat-value">{y_stats['min']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Y Max:</span> <span class="stat-value">{y_stats['max']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Y Mean:</span> <span class="stat-value">{y_stats['mean']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Y Median:</span> <span class="stat-value">{y_stats['median']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Y Std Dev:</span> <span class="stat-value">{y_stats['std']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Data Points:</span> <span class="stat-value">{y_stats['count']:,}</span></div>
    </div>
    '''
    
    z_stats_html = f'''
    <div class="statistics-box">
        <div class="stat-item"><span class="stat-label">Z Min:</span> <span class="stat-value">{z_stats['min']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Z Max:</span> <span class="stat-value">{z_stats['max']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Z Mean:</span> <span class="stat-value">{z_stats['mean']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Z Median:</span> <span class="stat-value">{z_stats['median']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Z Std Dev:</span> <span class="stat-value">{z_stats['std']:.4f} µm</span></div>
        <div class="stat-item"><span class="stat-label">Data Points:</span> <span class="stat-value">{z_stats['count']:,}</span></div>
    </div>
    '''
    
    # 創建完整的 HTML 與標籤頁
    html = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AutoZ Wafer4P Aligner - {machine_type}</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <script>
        function showTab(tabName) {{
            // 隱藏所有標籤頁內容
            var tabContents = document.getElementsByClassName("tab-content");
            for (var i = 0; i < tabContents.length; i++) {{
                tabContents[i].style.display = "none";
            }}
            
            // 移除所有標籤按鈕的活動類
            var tabButtons = document.getElementsByClassName("tab-button");
            for (var i = 0; i < tabButtons.length; i++) {{
                tabButtons[i].classList.remove("active");
            }}
            
            // 顯示選定的標籤頁內容並將按鈕標記為活動
            document.getElementById(tabName).style.display = "block";
            document.getElementById("btn-" + tabName).classList.add("active");
        }}
        </script>
        <style>
            body {{
                font-family: "Microsoft JhengHei", Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #FFFFFF;
            }}
            .container {{
                width: 95%;
                margin: 20px auto;
            }}
            .header {{
                background-color: #2D2D2D;
                color: #E0E0E0;
                padding: 15px 15px 10px 15px;
                text-align: center;
                border-radius: 8px 8px 0 0;
                margin-bottom: 0;
            }}
            .timestamp {{
                font-size: 12px;
                color: #AAAAAA;
                text-align: right;
                padding: 0 20px 10px 0;
                margin: 0;
                background-color: #2D2D2D;
            }}
            .tab-buttons {{
                display: flex;
                justify-content: center;
                padding: 10px;
                background-color: #333;
                border-bottom-left-radius: 8px;
                border-bottom-right-radius: 8px;
            }}
            .tab-button {{
                background: #454545;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                margin: 0 10px;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
                transition: background 0.3s;
            }}
            .tab-button:hover {{
                background: #555;
            }}
            .tab-button.active {{
                background: #93aec1;
            }}
            .tab-content {{
                display: none;
                margin-top: 20px;
            }}
            .chart-container {{
                margin: 20px auto;
                padding: 15px;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                width: 100%;
                max-width: 1200px;
                display: flex;
                flex-direction: column;
                align-items: center;
            }}
            .js-plotly-plot, .plot-container {{
                width: 100% !important;
                max-width: 1100px !important;
                margin: 0 auto !important;
            }}
            .footer {{
                text-align: center;
                font-size: 12px;
                color: #888;
                padding: 20px;
                margin-top: 20px;
                border-top: 1px solid #ddd;
            }}
            .statistics-container {{
                margin: 20px auto;
                padding: 15px;
                background-color: #f8f9fa;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                width: 100%;
                max-width: 1200px;
            }}
            .statistics-box {{
                display: flex;
                flex-wrap: wrap;
                justify-content: space-around;
                padding: 10px;
            }}
            .stat-item {{
                padding: 10px;
                margin: 5px;
                background-color: white;
                border-radius: 5px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                min-width: 150px;
                text-align: center;
            }}
            .stat-label {{
                font-weight: bold;
                display: block;
                margin-bottom: 5px;
                color: #555;
            }}
            .stat-value {{
                font-size: 18px;
                color: #333;
            }}
            .stats-title {{
                text-align: center;
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 10px;
                color: #333;
            }}
            .section-divider {{
                border-top: 1px solid #ddd;
                margin: 30px 0;
                width: 100%;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>AutoZ Wafer4P Aligner - {machine_type}</h1>
                <p>Data Visualization Analytics Tool</p>
            </div>
            
            <div class="timestamp">
                Generated on: {timestamp}
            </div>
            
            <div class="tab-buttons">
                <button id="btn-wafer-status" class="tab-button" onclick="showTab('wafer-status')">Wafer Status</button>
                <button id="btn-z-tab" class="tab-button active" onclick="showTab('z-tab')">Z Value</button>
                <button id="btn-x-tab" class="tab-button" onclick="showTab('x-tab')">X Value</button>
                <button id="btn-y-tab" class="tab-button" onclick="showTab('y-tab')">Y Value</button>
            </div>
            
            <!-- Wafer Status Dashboard Tab -->
            <div id="wafer-status" class="tab-content">
                {wafer_status_html}
            </div>
            
            <!-- Z Tab Content -->
            <div id="z-tab" class="tab-content">
                <!-- 1. Z Data Statistics (置頂) -->
                <div class="statistics-container">
                    <div class="stats-title">Z Data Statistics</div>
                    {z_stats_html}
                </div>
                
                <!-- 2. Z AutoZ Values 圖表 -->
                <div class="chart-container">
                    {z_html}
                </div>
                
                <div class="section-divider"></div>
                
                <!-- 3. Z Value Anomaly Analysis 圖表 (不顯示統計數據) -->
                <div class="chart-container">
                    <div class="chart-title" style="font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #333;">
                        Z Value Anomaly Analysis
                    </div>
                    {z_anomaly_html}
                </div>
            </div>
            
            <!-- X Tab Content -->
            <div id="x-tab" class="tab-content">
                <div class="statistics-container">
                    <div class="stats-title">X Data Statistics</div>
                    {x_stats_html}
                </div>
                
                <div class="chart-container">
                    {x_html}
                </div>
            </div>
            
            <!-- Y Tab Content -->
            <div id="y-tab" class="tab-content">
                <div class="statistics-container">
                    <div class="stats-title">Y Data Statistics</div>
                    {y_stats_html}
                </div>
                
                <div class="chart-container">
                    {y_html}
                </div>
            </div>
            
            <div class="footer">
                AutoZ Wafer4P Aligner | Document automatically generated
            </div>
        </div>
        
        <script>
            // 頁面載入時默認顯示 Z 標籤頁
            document.addEventListener('DOMContentLoaded', function() {{
                showTab('z-tab');
            }});
        </script>
    </body>
    </html>
    '''
    
    return html

# 定義工作線程類 - 處理AutoZLog.txt
class AutoZLogWorker(QThread):
    finished = pyqtSignal(object)  # 完成信號
    error = pyqtSignal(str)        # 錯誤信號
    progress = pyqtSignal(int)     # 進度信號

    def __init__(self, file_path, processor):
        super().__init__()
        self.file_path = file_path
        self.processor = processor

    def run(self):
        try:
            self.progress.emit(10)
            # 處理檔案
            timestamp = self.processor.process_autoz_log(self.file_path)
            self.progress.emit(100)
            # 發送完成信號
            self.finished.emit(timestamp)
        except Exception as e:
            # 發送錯誤信號
            self.error.emit(str(e))

# 定義工作線程類 - 處理ALL.TXT
class AllTxtWorker(QThread):
    finished = pyqtSignal(object)  # 完成信號
    error = pyqtSignal(str)        # 錯誤信號
    progress = pyqtSignal(int)     # 進度信號

    def __init__(self, file_path, processor, timestamp):
        super().__init__()
        self.file_path = file_path
        self.processor = processor
        self.timestamp = timestamp

    def run(self):
        try:
            self.progress.emit(10)
            # 處理檔案
            result = self.processor.process_all_txt(self.file_path, self.timestamp)
            self.progress.emit(90)
            # 發送完成信號
            self.finished.emit(result)
        except Exception as e:
            # 發送錯誤信號
            self.error.emit(str(e))

# 定義工作線程類 - 生成HTML
class HtmlGeneratorWorker(QThread):
    finished = pyqtSignal(str)     # 完成信號，傳回HTML文件路徑
    error = pyqtSignal(str)        # 錯誤信號
    progress = pyqtSignal(int)     # 進度信號

    def __init__(self, wafer_data, x_standard, y_standard, z_standard, machine_type):
        super().__init__()
        self.wafer_data = wafer_data
        self.x_standard = x_standard
        self.y_standard = y_standard
        self.z_standard = z_standard
        self.machine_type = machine_type

    def run(self):
        try:
            self.progress.emit(10)
            # 生成HTML內容
            html_content = generate_html_report(
                self.wafer_data, 
                self.x_standard,
                self.y_standard,
                self.z_standard,
                self.machine_type
            )
            
            self.progress.emit(50)
            
            # 創建臨時文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_file_path = f.name
            
            self.progress.emit(90)
            # 發送完成信號
            self.finished.emit(temp_file_path)
        except Exception as e:
            # 發送錯誤信號
            self.error.emit(str(e))

class DataVisualizer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.check_version()
        self.save_log()
        
        # 設定窗口屬性
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.WindowCloseButtonHint)      
        self.setStyleSheet(STYLE_SHEET)
        
        # 設定窗口標題和大小
        self.setWindowTitle('AutoZ Wafer4P Aligner')
        self.setGeometry(100, 100, 1300, 900)
        
        # 創建主佈局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        
        # 創建左側控制面板
        control_panel = QWidget()
        control_layout = QVBoxLayout(control_panel)
        control_panel.setFixedWidth(300)
        
        # 添加控制元素到左側面板
        self.create_control_panel(control_layout)
        
        # 添加進度條
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        control_layout.addWidget(self.progress)
        
        # 添加彈性空間
        control_layout.addStretch()
        
        # 創建右側圖形顯示區域
        self.web_view = QWebEngineView()
        self.web_view.setMinimumSize(1300, 950)
        
        # 添加左側控制面板和右側圖形顯示到主佈局
        main_layout.addWidget(control_panel)
        main_layout.addWidget(self.web_view)
        
        # 初始化變數
        self.AutoZLog_Last_Trigger4pinalignment_Time = None
        self.ALL_AutoZ_X_Stardand = None
        self.ALL_AutoZ_Y_Stardand = None
        self.ALL_AutoZ_Z_Stardand = None
        self.wafer_data = {}  # 存儲晶圓數據的字典
        
        # 初始化處理器變數
        self.processor = None
        self.current_machine_type = None  # 儲存當前機台類型
        
        # 初始化工作線程變數
        self.autoz_worker = None
        self.all_txt_worker = None
        self.html_worker = None
        
        # 顯示默認歡迎頁面
        self.show_default_plot()

    def create_control_panel(self, layout):
        """創建控制面板"""
        control_layout = QVBoxLayout()
        
        # 添加機台選擇下拉選單
        machine_layout = QHBoxLayout()
        machine_label = QLabel("Machine Type")
        self.machine_combo = QComboBox()
        self.machine_combo.addItem("") # 添加空白選項作為預設
        self.machine_combo.addItems(['J750', 'J750EX', 'UFLEX', 'ETS88', 'Accotest', 'AG93000', 'T2K'])
        self.machine_combo.setCurrentIndex(0) # 設置空白選項為當前選擇
        self.machine_combo.setPlaceholderText("Select Machine") # 設置佔位符文字
        self.machine_combo.currentTextChanged.connect(self.on_machine_changed)
        machine_layout.addWidget(machine_label)
        machine_layout.addWidget(self.machine_combo)
        control_layout.addLayout(machine_layout)
        
        # 添加間隔
        spacer = QLabel("")
        spacer.setFixedHeight(10)
        control_layout.addWidget(spacer)
        
        # AutoZLog.txt 按鈕
        self.select_file_button = QPushButton('Select AutoZLog.txt')
        self.select_file_button.clicked.connect(self.select_file)
        control_layout.addWidget(self.select_file_button)
        
        # ALL.txt 按鈕
        self.select_all_txt_button = QPushButton('Select ALL.txt')
        self.select_all_txt_button.clicked.connect(self.select_all_txt)
        control_layout.addWidget(self.select_all_txt_button)
        
        layout.addLayout(control_layout)

    def on_machine_changed(self, machine_type):
        """當機台類型變更時調用"""
        print(f"Selected machine type {machine_type}")
        
        # 重置按鈕狀態
        self.select_file_button.setEnabled(True)
        self.select_all_txt_button.setEnabled(True)
        
        # 重置數據
        self.AutoZLog_Last_Trigger4pinalignment_Time = None
        self.ALL_AutoZ_X_Stardand = None
        self.ALL_AutoZ_Y_Stardand = None
        self.ALL_AutoZ_Z_Stardand = None
        self.wafer_data = {}
        
        # 儲存當前機台類型
        self.current_machine_type = machine_type
        
        if machine_type in ['J750', 'J750EX', 'UFLEX']:
            self.processor = J750_J750EX_UFLEX_process_V3
        elif machine_type in ['ETS88', 'Accotest']:
            self.processor = ETS88_Accotest_process_V3
        elif machine_type == 'AG93000':
            self.processor = AG93000_process_V4
        elif machine_type == 'T2K':
            self.processor = T2K_process_V1
        else:
            self.processor = None
            self.current_machine_type = None
        
        # 重新顯示默認頁面
        self.show_default_plot()

    def select_file(self):
        """處理檔案選擇、驗證和處理"""
        if not self.processor:
            QMessageBox.warning(self, "Warning", "Please select machine type first!")
            return
            
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select AutoZ File",
            "",
            "All Files (*.*);;Text and Log Files (*.txt *.log);;Text Files (*.txt);;Log Files (*.log)"
        )
        
        if file_path:
            # 顯示進度條
            self.progress.setVisible(True)
            self.progress.setValue(0)
            
            # 創建並設置工作線程
            self.autoz_worker = AutoZLogWorker(file_path, self.processor)
            
            # 連接信號
            self.autoz_worker.finished.connect(self.on_autoz_finished)
            self.autoz_worker.error.connect(self.on_worker_error)
            self.autoz_worker.progress.connect(self.progress.setValue)
            
            # 禁用按鈕防止重複點擊
            self.select_file_button.setEnabled(False)
            
            # 啟動線程
            self.autoz_worker.start()

    def on_autoz_finished(self, timestamp):
        """AutoZLog處理完成的回調"""
        self.AutoZLog_Last_Trigger4pinalignment_Time = timestamp
        self.progress.setVisible(False)
        print(f"Processing completed, timestamp: {self.AutoZLog_Last_Trigger4pinalignment_Time}")
        # 保持按鈕禁用狀態
        self.select_file_button.setEnabled(False)

    def select_all_txt(self):
        """處理 ALL.TXT 檔案選擇、驗證和處理"""
        if not self.processor:
            QMessageBox.warning(self, "Warning", "Please select machine type first!")
            return
            
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select ALL File",
            "",
            "All Files (*.*);;Text and Log Files (*.txt *.log);;Text Files (*.txt);;Log Files (*.log)"
        )
        
        if file_path:
            # 檢查 AutoZLog 時間戳是否可用
            if not self.AutoZLog_Last_Trigger4pinalignment_Time:
                error_msg = QMessageBox()
                error_msg.setIcon(QMessageBox.Warning)
                error_msg.setWindowTitle("Processing Warning")
                error_msg.setText("AutoZLog timestamp not available!")
                error_msg.setInformativeText("Please process AutoZLog.txt file first.")
                error_msg.exec_()
                return
            
            # 顯示進度條
            self.progress.setVisible(True)
            self.progress.setValue(0)
            
            # 創建並設置工作線程
            self.all_txt_worker = AllTxtWorker(
                file_path, 
                self.processor, 
                self.AutoZLog_Last_Trigger4pinalignment_Time
            )
            
            # 連接信號
            self.all_txt_worker.finished.connect(self.on_all_txt_finished)
            self.all_txt_worker.error.connect(self.on_worker_error)
            self.all_txt_worker.progress.connect(self.progress.setValue)
            
            # 禁用按鈕防止重複點擊
            self.select_all_txt_button.setEnabled(False)
            
            # 啟動線程
            self.all_txt_worker.start()

    def on_all_txt_finished(self, result):
        """ALL.TXT處理完成的回調"""
        self.wafer_data = result['wafer_data']
        self.ALL_AutoZ_X_Stardand = result['x_standard']
        self.ALL_AutoZ_Y_Stardand = result['y_standard']
        self.ALL_AutoZ_Z_Stardand = result['z_standard']
        
        print(f"Successfully processed ALL.TXT file")
        
        # 處理完成後自動生成 HTML 文件
        self.generate_html_file()

    def on_worker_error(self, error_msg):
        """工作線程錯誤處理"""
        self.progress.setVisible(False)
        
        # 重新啟用按鈕
        self.select_file_button.setEnabled(True)
        self.select_all_txt_button.setEnabled(True)
        
        # 顯示錯誤訊息
        error_msg_box = QMessageBox()
        error_msg_box.setIcon(QMessageBox.Critical)
        error_msg_box.setWindowTitle("Processing Error")
        error_msg_box.setText("Error processing file!")
        error_msg_box.setInformativeText(f"Details: {error_msg}")
        error_msg_box.exec_()
        print(f"Error processing file: {error_msg}")

    def generate_html_file(self):
        """生成 HTML 檔案並在瀏覽器中顯示"""
        try:
            # 顯示進度條
            self.progress.setVisible(True)
            self.progress.setValue(0)
            
            # 檢查是否有數據
            if not self.wafer_data:
                QMessageBox.warning(self, 'No Data', 'No wafer data available for visualization.')
                self.progress.setVisible(False)
                return
            
            # 創建並設置HTML生成工作線程
            self.html_worker = HtmlGeneratorWorker(
                self.wafer_data, 
                self.ALL_AutoZ_X_Stardand,
                self.ALL_AutoZ_Y_Stardand,
                self.ALL_AutoZ_Z_Stardand,
                self.current_machine_type  # 傳遞機台類型
            )
            
            # 連接信號
            self.html_worker.finished.connect(self.on_html_generated)
            self.html_worker.error.connect(self.on_worker_error)
            self.html_worker.progress.connect(self.progress.setValue)
            
            # 啟動線程
            self.html_worker.start()
            
        except Exception as e:
            self.progress.setVisible(False)
            QMessageBox.critical(self, 'Error', f'Error generating HTML file:\n{str(e)}')
            print(f"Error in generate_html_file: {str(e)}")

    def on_html_generated(self, temp_file_path):
        """HTML生成完成的回調"""
        self.progress.setValue(100)
        QApplication.processEvents()
        
        # 在瀏覽器中打開
        webbrowser.open('file://' + os.path.abspath(temp_file_path))
        
        # 最小化應用程序窗口
        self.hide()
        
        self.progress.setVisible(False)
        
        # 直接關閉程式
        print("HTML report generated and opened. Closing application...")
        QApplication.quit()

    def show_default_plot(self):
        """顯示歡迎頁面"""
        try:
            welcome_html = '''
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {
                        margin: 0;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        background-color: #F8F9FA;
                        font-family: "微軟正黑體", sans-serif;
                        font-weight: bold;
                    }
                    .welcome-container {
                        text-align: center;
                        padding: 20px;
                    }
                    .title {
                        font-size: 36px;
                        color: #333;
                        margin-bottom: 20px;
                    }
                    .title-eng {
                        font-size: 32px;
                        color: #333;
                        margin-bottom: 20px;
                    }
                    .subtitle {
                        font-size: 20px;
                        color: #666;
                        margin-bottom: 30px;
                    }
                    .instruction {
                        font-size: 16px;
                        color: #444;
                        line-height: 1.6;
                    }
                </style>
            </head>
            <body>
                <div class="welcome-container">
                    <div class="title">Welcome</div>
                    <div class="title-eng">AutoZ Wafer4P Aligner</div>
                    <div class="subtitle">Data Visualization Analytics Tool</div>
                    <div class="instruction">
                        Please select machine type first, then upload required files to begin.
                    </div>
                </div>
            </body>
            </html>
            '''
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(welcome_html)
                self.web_view.setUrl(QUrl.fromLocalFile(f.name))

        except Exception as e:
            print(f"Error displaying welcome page: {str(e)}")

    def check_version(self):
        """檢查版本"""
        try:
            app_folder = os.path.normpath(os.path.join("M:", "BI_Database", "Apps", "Database", "Apps_Installation_package", "RD_All"))
            exe_files = [os.path.join(app_folder, f) for f in os.listdir(app_folder) 
                        if f.startswith("AutoZ Wafer4P Aligner_V") and f.endswith(".exe")]

            if not exe_files:
                msg = QMessageBox()
                msg.setWindowTitle('No Launch Permission')
                msg.setText('Failed to get launch permission. Please apply for M:\\BI_Database authorization and contact DA Team')
                msg.setIcon(QMessageBox.Warning)
                screen = QApplication.primaryScreen().geometry()
                x = (screen.width() - msg.width()) // 2
                y = (screen.height() - msg.height()) // 2
                msg.move(x, y)
                msg.exec_()
                sys.exit(1)

            def parse_version(version_str):
                match = re.search(r'_V(\d+)\.(\d+)', version_str)
                if match:
                    return tuple(map(int, match.groups()))
                return (0, 0)

            latest_version = max(parse_version(os.path.basename(f)) for f in exe_files)
            latest_exe = max((f for f in exe_files 
                            if parse_version(os.path.basename(f)) == latest_version), 
                        key=os.path.getmtime)

            current_version_match = re.search(r'_V(\d+)\.(\d+)', os.path.basename(sys.executable))
            if current_version_match:
                current_version = tuple(map(int, current_version_match.groups()))
            else:
                current_version = (12, 0)  # 更新為新版本號

            if current_version[0] != latest_version[0]:
                msg = QMessageBox()
                msg.setWindowTitle('Update Notification')
                msg.setText(f'Updating to new version V{latest_version[0]}.{latest_version[1]}')
                msg.setIcon(QMessageBox.Information)
                screen = QApplication.primaryScreen().geometry()
                x = (screen.width() - msg.width()) // 2
                y = (screen.height() - msg.height()) // 2
                msg.move(x, y)
                msg.exec_()

                bat_content = '''@echo off
    timeout /t 2 /nobreak
    rmdir /s /q "C:\\Users\\{username}\\BITools\\AutoZ Wafer4P Aligner"
    del "%~f0"
    '''
                bat_path = os.path.join(os.environ['TEMP'], 'delete_autoz_wafer4p_aligner.bat')
                with open(bat_path, 'w') as f:
                    f.write(bat_content)
                
                subprocess.Popen(['cmd', '/c', bat_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                os.startfile(latest_exe)
                sys.exit(1)

            if current_version[0] != 12:
                msg = QMessageBox()
                msg.setWindowTitle('Version Error')
                msg.setText('Software version error')
                msg.setIcon(QMessageBox.Warning)
                screen = QApplication.primaryScreen().geometry()
                x = (screen.width() - msg.width()) // 2
                y = (screen.height() - msg.height()) // 2
                msg.move(x, y)
                msg.exec_()
                sys.exit(1)

        except FileNotFoundError:
            msg = QMessageBox()
            msg.setWindowTitle('No Launch Permission')
            msg.setText('Failed to get launch permission. Please apply for M:\\BI_Database authorization and contact DA Team')
            msg.setIcon(QMessageBox.Warning)
            screen = QApplication.primaryScreen().geometry()
            x = (screen.width() - msg.width()) // 2
            y = (screen.height() - msg.height()) // 2
            msg.move(x, y)
            msg.exec_()
            sys.exit(1)

    def save_log(self):
        """保存日誌到資料庫"""
        try:
            current_datetime = datetime.now()
            try:
                conn_str = f'DRIVER={{SQL Server}};SERVER={SQL_SERVER_INFO["server"]};DATABASE={SQL_SERVER_INFO["database"]};UID={SQL_SERVER_INFO["username"]};PWD={SQL_SERVER_INFO["password"]};App=AutoZ Wafer4P Aligner'
                with pyodbc.connect(conn_str) as conn:
                    cursor = conn.cursor()

                    insert_query = f"""
                    INSERT INTO {SQL_SERVER_INFO["apps_log_table"]} (Activation_Time, User_Id, Status, Apps_Name)
                    VALUES (?, ?, ?, ?)
                    """
                    cursor.execute(insert_query, (current_datetime, username, "Open", "AutoZ Wafer4P Aligner"))
                    conn.commit()
            except pyodbc.Error as e:
                print(f"SQL Server log database connection error: {str(e)}")

            except Exception as e:
                print(f"Log database operation error: {str(e)}")

        except Exception as e:
            print(f"Error occurred while writing to log database: {e}")
            
    def showEvent(self, event):
        """窗口顯示時觸發，使其居中"""
        super().showEvent(event)
        # 確保只在第一次顯示時居中
        if not hasattr(self, '_window_centered'):
            self.center_window()
            self._window_centered = True

    def center_window(self):
        """使窗口居中顯示"""
        # 獲取主屏幕尺寸
        screen = QApplication.primaryScreen().availableGeometry()
        # 獲取窗口尺寸
        window_size = self.frameGeometry()
        # 計算中心點
        center_point = screen.center()
        # 移動窗口
        window_size.moveCenter(center_point)
        self.move(window_size.topLeft())
        
    def closeEvent(self, event):
        """關閉窗口時執行"""
        try:
            # 停止所有運行中的線程
            if hasattr(self, 'autoz_worker') and self.autoz_worker and self.autoz_worker.isRunning():
                self.autoz_worker.terminate()
                self.autoz_worker.wait()
                
            if hasattr(self, 'all_txt_worker') and self.all_txt_worker and self.all_txt_worker.isRunning():
                self.all_txt_worker.terminate()
                self.all_txt_worker.wait()
                
            if hasattr(self, 'html_worker') and self.html_worker and self.html_worker.isRunning():
                self.html_worker.terminate()
                self.html_worker.wait()
                
            # 執行原始關閉事件
            super().closeEvent(event)
            
        except Exception as e:
            print(f"Error during window close: {str(e)}")
            super().closeEvent(event)

def main():
    app = QApplication(sys.argv)
    
    # 設置字體
    font = QFont("微軟正黑體", 9)
    font.setBold(True)
    app.setFont(font)
    
    # 設置應用程序樣式表
    app.setStyleSheet(STYLE_SHEET)
    
    # 設置圖標
    application_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(application_path, 'format.ico')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    # 顯示主窗口
    window = DataVisualizer()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()