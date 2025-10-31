# AutoZ Wafer4P Aligner V12.0
# Data Visualization Analytics Tool

import sys
import os
import pandas as pd
import re
import numpy as np
import math
import time
from datetime import datetime
from flask import Flask, jsonify, request, redirect
from tkinter import Tk, filedialog
import socket
import threading
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

# 取得使用者名稱
username = os.environ.get('USERNAME', 'Unknown')

# 從 JSON 檔案讀取 SQL Server 連線資訊
json_path = r"M:\BI_Database\Apps\Database\Apps_Database\O_All\SQL_Server\SQL_Server_Info_User_BI.json"
with open(json_path, 'r') as file:
    sql_connection_info = json.load(file)

# SQL Server 連線資訊
SQL_SERVER_INFO = {
    "server": sql_connection_info["server"],
    "database": sql_connection_info["database"], 
    "username": sql_connection_info["username"],
    "password": sql_connection_info["password"],
    "apps_log_table": sql_connection_info["apps_log_table"]  
}

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ==================== Flask 應用程式初始化 ====================
app = Flask(__name__)

# ==================== 全域變數 ====================
current_worker = None
last_activity_time = time.time()
ACTIVITY_TIMEOUT = 30 * 60  # 30 分鐘無活動自動關閉

# 分析相關全域變數
selected_machine_type = None
autoz_log_timestamp = None
analysis_file_data = None
processor_module = None


# ==================== 工具函數 ====================

def update_activity():
    """更新最後活動時間"""
    global last_activity_time
    last_activity_time = time.time()


def check_activity_thread():
    """背景執行緒：檢查活動超時,超過 30 分鐘無活動則關閉應用程式"""
    global last_activity_time
    
    while True:
        time.sleep(60)
        elapsed = time.time() - last_activity_time
        
        if elapsed > ACTIVITY_TIMEOUT:
            print(f"Activity timeout reached ({ACTIVITY_TIMEOUT/60} minutes). Shutting down...")
            os._exit(0)


def is_port_in_use(port):
    """檢查指定埠是否被使用"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0


def find_available_port(start_port=8000, end_port=8999):
    """尋找可用的埠號 (範圍 8000-8999)"""
    for port in range(start_port, end_port + 1):
        if not is_port_in_use(port):
            return port
    raise RuntimeError(f"No available ports in range {start_port}-{end_port}")


def check_version():
    """檢查版本更新"""
    try:
        app_folder = os.path.normpath(os.path.join("M:", "BI_Database", "Apps", "Database", "Apps_Installation_package", "RD_All"))
        exe_files = [os.path.join(app_folder, f) for f in os.listdir(app_folder) 
                    if f.startswith("AutoZ Wafer4P Aligner_V") and f.endswith(".exe")]

        if not exe_files:
            return {
                'status': 'error',
                'type': 'permission',
                'message': 'Failed to get launch permission.'
            }

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
            current_version = (12, 0)

        if current_version[0] != latest_version[0]:
            return {
                'status': 'update',
                'message': f'Will update to new version V{latest_version[0]}.{latest_version[1]}',
                'latest_exe': latest_exe
            }

        if current_version[0] != 12:
            return {
                'status': 'error',
                'type': 'version',
                'message': 'Software version error'
            }

        return {'status': 'ok'}

    except FileNotFoundError:
        return {
            'status': 'error',
            'type': 'permission',
            'message': 'Failed to get launch permission.'
        }


def save_log():
    """儲存使用記錄到 SQL Server"""
    try:
        current_datetime = datetime.now()
        
        conn_str = (
            f'DRIVER={{SQL Server}};'
            f'SERVER={SQL_SERVER_INFO["server"]};'
            f'DATABASE={SQL_SERVER_INFO["database"]};'
            f'UID={SQL_SERVER_INFO["username"]};'
            f'PWD={SQL_SERVER_INFO["password"]};'
            f'App=AutoZ Wafer4P Aligner'
        )
        
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            
            insert_query = f"""
            INSERT INTO {SQL_SERVER_INFO["apps_log_table"]} 
            (Activation_Time, User_Id, Status, Apps_Name)
            VALUES (?, ?, ?, ?)
            """
            
            cursor.execute(insert_query, 
                          (current_datetime, username, "Open", "AutoZ Wafer4P Aligner"))
            conn.commit()
            
            print("Log saved successfully")
            
    except pyodbc.Error as e:
        print(f"SQL Server connection error: {str(e)}")
    except Exception as e:
        print(f"Error saving log: {str(e)}")


# ==================== 圖表生成函數 ====================

def create_line_chart(wafer_data, axis_type, standard_value=None, standard_point_data=None):
    """為指定的軸類型創建折線圖
    
    Args:
        wafer_data: 晶圓資料字典
        axis_type: 軸類型 ('x', 'y', 'z')
        standard_value: 標準參考值 (僅用於 Z 軸)
        standard_point_data: AutoZ complete 點位資料
        
    Returns:
        tuple: (fig, stats) Plotly 圖表物件和統計資料
    """
    fig = go.Figure()
    
    # 追蹤最大/最小值用於註釋定位
    max_y_value = float('-inf')
    min_y_value = float('inf')
    
    # 為 x 軸創建連續索引
    continuous_x = []
    continuous_y = []
    wafer_boundaries = []
    wafer_labels = []
    point_labels = []
    
    value_key = f"{axis_type}_values"
    
    # 按開始時間排序晶圓
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    
    current_index = 0
    
    # 處理每個晶圓
    for wafer_id, data in sorted_wafers:
        if value_key in data and data[value_key]:
            values = data[value_key]
            
            start_idx = current_index
            end_idx = start_idx + len(values)
            
            continuous_x.extend(range(start_idx, end_idx))
            continuous_y.extend(values)
            point_labels.extend([wafer_id] * len(values))
            
            wafer_boundaries.append((start_idx, end_idx - 1))
            wafer_labels.append(wafer_id)
            
            current_index = end_idx
            
            if values:
                max_y_value = max(max_y_value, max(values))
                min_y_value = min(min_y_value, min(values))
    
    # 準備顏色和大小陣列
    if continuous_x and continuous_y:
        colors = []
        sizes = []
        
        # 檢查第一個點是否為 Auto Z complete 點
        is_first_point_autoz = False
        if standard_point_data and axis_type in standard_point_data:
            autoz_value = standard_point_data[axis_type]
            if len(continuous_y) > 0 and abs(continuous_y[0] - autoz_value) < 0.001:
                is_first_point_autoz = True
                max_y_value = max(max_y_value, autoz_value)
                min_y_value = min(min_y_value, autoz_value)
        
        for i in range(len(continuous_x)):
            if i == 0 and is_first_point_autoz:
                colors.append('#e5857b')
                sizes.append(20)
                point_labels[0] = "AutoZ Complete"
            else:
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
        
        # 在 Z 標準線上添加常駐標籤
        if continuous_x:
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
    
    # 添加晶圓邊界標記
    for i, (start, end) in enumerate(wafer_boundaries):
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
    
    # 計算統計數據
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
    """創建 Wafer 狀態儀表板,顯示哪些 Wafer 有低於標準的 Z 值
    
    Args:
        wafer_data: 晶圓資料字典
        z_standard: Z 標準值
        
    Returns:
        str: Wafer 狀態儀表板的 HTML 內容
    """
    wafer_status = {}
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    
    for wafer_id, data in sorted_wafers:
        if 'z_values' in data and data['z_values']:
            below_standard = any(z < z_standard for z in data['z_values'])
            below_count = sum(1 for z in data['z_values'] if z < z_standard)
            total_count = len(data['z_values'])
            percent_below = (below_count / total_count * 100) if total_count > 0 else 0
            
            wafer_status[wafer_id] = {
                'below_standard': below_standard,
                'below_count': below_count,
                'total_count': total_count,
                'percent_below': percent_below
            }
    
    wafers_per_row = 4
    
    html_content = '''
    <div class="dashboard-container">
        <h2 class="dashboard-title">Wafer Status Dashboard</h2>
        <p class="dashboard-description">Status shows wafers with Z values below standard (red = has points below standard)</p>
        <div class="wafer-grid">
    '''
    
    row_count = math.ceil(len(wafer_status) / wafers_per_row)
    wafer_count = 0
    
    for wafer_id, status in wafer_status.items():
        card_class = "wafer-card-red" if status['below_standard'] else "wafer-card-green"
        
        if wafer_count % wafers_per_row == 0:
            html_content += '<div class="wafer-row">'
        
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
        
        if wafer_count % wafers_per_row == 0 or wafer_count == len(wafer_status):
            html_content += '</div>'
    
    html_content += '''
        </div>
    </div>
    '''
    
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
    """創建突顯低於標準的 Z 值的圖表
    
    Args:
        wafer_data: 晶圓資料字典
        z_standard: Z 標準值
        standard_point_data: AutoZ complete 點位資料
        
    Returns:
        tuple: (fig, stats) Plotly 圖表物件和統計資料
    """
    fig = go.Figure()
    
    continuous_x = []
    normal_y = []
    anomaly_y = []
    normal_indices = []
    anomaly_indices = []
    wafer_ids = []
    
    sorted_wafers = sorted(wafer_data.items(), key=lambda x: x[1]['start_time'])
    current_index = 0
    
    for wafer_id, data in sorted_wafers:
        if 'z_values' in data and data['z_values']:
            values = data['z_values']
            indices = list(range(current_index, current_index + len(values)))
            
            for idx, z_val in zip(indices, values):
                wafer_ids.append(wafer_id)
                if z_val < z_standard:
                    anomaly_indices.append(idx)
                    anomaly_y.append(z_val)
                else:
                    normal_indices.append(idx)
                    normal_y.append(z_val)
            
            current_index += len(values)
    
    # 檢查第一個點是否為 Auto Z complete 點
    is_first_point_autoz = False
    if standard_point_data and 'z' in standard_point_data:
        autoz_z_value = standard_point_data['z']
        if len(wafer_ids) > 0:
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
        
        normal_colors = []
        normal_sizes = []
        normal_symbols = []
        
        for i, idx in enumerate(normal_indices):
            if idx == 0 and is_first_point_autoz:
                normal_colors.append('#FF6600')
                normal_sizes.append(14)
                normal_symbols.append('diamond')
            else:
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
        
        anomaly_colors = []
        anomaly_sizes = []
        anomaly_symbols = []
        
        for i, idx in enumerate(anomaly_indices):
            if idx == 0 and is_first_point_autoz:
                anomaly_colors.append('#FF6600')
                anomaly_sizes.append(14)
                anomaly_symbols.append('diamond')
            else:
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
    
    stats = {
        'total_points': total_points,
        'normal_points': len(normal_indices),
        'anomaly_points': anomaly_count,
        'anomaly_percent': anomaly_percent
    }
    
    return fig, stats


# ==================== Worker 函數 ====================

def process_autoz_log_worker(file_path):
    """處理 AutoZLog.txt 檔案"""
    try:
        timestamp = processor_module.process_autoz_log(file_path)
        return {'success': True, 'timestamp': timestamp, 'error': None}
    except Exception as e:
        return {'success': False, 'timestamp': None, 'error': str(e)}


def process_all_txt_worker(file_path, timestamp):
    """處理 ALL.txt 檔案"""
    try:
        result = processor_module.process_all_txt(file_path, timestamp)
        return {'success': True, 'result': result, 'error': None}
    except Exception as e:
        return {'success': False, 'result': None, 'error': str(e)}


# ==================== Flask 路由 ====================

@app.route('/')
def index():
    """主頁面路由"""
    update_activity()
    return generate_index_html()


@app.route('/result')
def result():
    """結果頁面路由"""
    update_activity()
    
    global analysis_file_data
    
    if analysis_file_data is None:
        return redirect('/')
    
    return generate_result_html(analysis_file_data)


# ==================== API 端點 ====================

@app.route('/api/heartbeat', methods=['POST'])
def heartbeat():
    """心跳端點,用於保持活動狀態"""
    update_activity()
    return jsonify({'status': 'ok'})


@app.route('/api/check_version', methods=['GET'])
def api_check_version():
    """版本檢查 API"""
    update_activity()
    result = check_version()
    return jsonify(result)


@app.route('/api/execute_update', methods=['POST'])
def execute_update():
    """執行更新：啟動新版本安裝程式並退出當前程式"""
    try:
        update_activity()
        
        data = request.get_json()
        latest_exe = data.get('latest_exe')
        
        if not latest_exe:
            return jsonify({'success': False, 'error': 'No latest_exe provided'})
        
        # 創建批次檔用於延遲刪除舊版本暫存目錄
        bat_content = '''@echo off
timeout /t 3 /nobreak
rmdir /s /q "C:\\Users\\{username}\\BITools\\AutoZ Wafer4P Aligner"
del "%~f0"
'''
        bat_path = os.path.join(os.environ['TEMP'], 'delete_AutoZ_Wafer4P_Aligner.bat')
        with open(bat_path, 'w') as f:
            f.write(bat_content)
        
        # 在背景執行批次檔
        subprocess.Popen(['cmd', '/c', bat_path], shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW)
        
        # 啟動新版本安裝程式
        os.startfile(latest_exe)
        
        # 給予短暫延遲後退出程式
        time.sleep(0.5)
        os._exit(0)
        
        return jsonify({'success': True})
        
    except Exception as e:
        print(f"Error executing update: {str(e)}")
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/select_machine', methods=['POST'])
def select_machine():
    """機台選擇 API"""
    update_activity()
    
    data = request.get_json()
    machine_type = data.get('machine_type', '')
    
    global selected_machine_type, processor_module, autoz_log_timestamp, analysis_file_data
    
    # 重置狀態
    selected_machine_type = None
    processor_module = None
    autoz_log_timestamp = None
    analysis_file_data = None
    
    # 根據機台類型載入對應的處理模組
    if machine_type in ['J750', 'J750EX', 'UFLEX']:
        processor_module = J750_J750EX_UFLEX_process_V3
        selected_machine_type = machine_type
    elif machine_type in ['ETS88', 'Accotest']:
        processor_module = ETS88_Accotest_process_V3
        selected_machine_type = machine_type
    elif machine_type == 'AG93000':
        processor_module = AG93000_process_V4
        selected_machine_type = machine_type
    elif machine_type == 'T2K':
        processor_module = T2K_process_V1
        selected_machine_type = machine_type
    else:
        return jsonify({
            'success': False,
            'error': 'Invalid machine type'
        })
    
    print(f"Machine type selected: {machine_type}")
    
    return jsonify({
        'success': True,
        'machine_type': machine_type
    })


@app.route('/api/select_file', methods=['POST'])
def select_file():
    """檔案選擇 API (使用 Tkinter 檔案對話框)"""
    update_activity()
    
    data = request.get_json()
    file_type = data.get('file_type', '')
    
    if not processor_module:
        return jsonify({
            'success': False,
            'error': 'Please select machine type first'
        })
    
    try:
        # 創建 Tkinter root 視窗 (隱藏)
        root = Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # 開啟檔案對話框
        file_path = filedialog.askopenfilename(
            title=f"Select {file_type} File",
            filetypes=[
                ("All Files", "*.*"),
                ("Text and Log Files", "*.txt *.log"),
                ("Text Files", "*.txt"),
                ("Log Files", "*.log")
            ]
        )
        
        root.destroy()
        
        if not file_path:
            return jsonify({
                'success': False,
                'error': 'No file selected'
            })
        
        print(f"File selected: {file_path}")
        
        return jsonify({
            'success': True,
            'file_path': file_path,
            'file_name': os.path.basename(file_path)
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


@app.route('/api/process_autoz_log', methods=['POST'])
def api_process_autoz_log():
    """處理 AutoZLog.txt API"""
    update_activity()
    
    global autoz_log_timestamp
    
    data = request.get_json()
    file_path = data.get('file_path', '')
    
    if not file_path:
        return jsonify({
            'success': False,
            'error': 'No file path provided'
        })
    
    if not processor_module:
        return jsonify({
            'success': False,
            'error': 'Please select machine type first'
        })
    
    print(f"Processing AutoZLog.txt: {file_path}")
    
    result = process_autoz_log_worker(file_path)
    
    if result['success']:
        autoz_log_timestamp = result['timestamp']
        print(f"AutoZLog processed successfully. Timestamp: {autoz_log_timestamp}")
        
        return jsonify({
            'success': True,
            'message': 'AutoZLog.txt processed successfully'
        })
    else:
        print(f"Error processing AutoZLog: {result['error']}")
        return jsonify({
            'success': False,
            'error': result['error']
        })


@app.route('/api/process_all_txt', methods=['POST'])
def api_process_all_txt():
    """處理 ALL.txt API"""
    update_activity()
    
    global analysis_file_data
    
    data = request.get_json()
    file_path = data.get('file_path', '')
    
    if not file_path:
        return jsonify({
            'success': False,
            'error': 'No file path provided'
        })
    
    if not processor_module:
        return jsonify({
            'success': False,
            'error': 'Please select machine type first'
        })
    
    if not autoz_log_timestamp:
        return jsonify({
            'success': False,
            'error': 'AutoZLog timestamp not available. Please process AutoZLog.txt first.'
        })
    
    print(f"Processing ALL.txt: {file_path}")
    
    result = process_all_txt_worker(file_path, autoz_log_timestamp)
    
    if result['success']:
        analysis_file_data = result['result']
        print("ALL.txt processed successfully")
        
        return jsonify({
            'success': True,
            'message': 'Analysis completed successfully',
            'redirect_url': '/result'
        })
    else:
        print(f"Error processing ALL.txt: {result['error']}")
        return jsonify({
            'success': False,
            'error': result['error']
        })


@app.route('/api/regenerate_chart', methods=['POST'])
def regenerate_chart():
    """重新生成圖表 API (用於 X/Y/Z 軸切換)"""
    update_activity()
    
    global analysis_file_data
    
    if analysis_file_data is None:
        return jsonify({
            'success': False,
            'error': 'No analysis data available'
        })
    
    data = request.get_json()
    axis_type = data.get('axis_type', 'z').lower()
    
    if axis_type not in ['x', 'y', 'z']:
        return jsonify({
            'success': False,
            'error': 'Invalid axis type'
        })
    
    try:
        wafer_data = analysis_file_data['wafer_data']
        x_standard = analysis_file_data['x_standard']
        y_standard = analysis_file_data['y_standard']
        z_standard = analysis_file_data['z_standard']
        
        standard_point_data = {
            'x': x_standard,
            'y': y_standard,
            'z': z_standard
        }
        
        # 根據軸類型決定是否傳入標準值
        standard_value = z_standard if axis_type == 'z' else None
        
        fig, stats = create_line_chart(
            wafer_data, 
            axis_type, 
            standard_value,
            standard_point_data
        )
        
        # 將圖表轉為字典格式
        chart_dict = fig.to_dict()
        
        return jsonify({
            'success': True,
            'chart': chart_dict,
            'stats': stats
        })
        
    except Exception as e:
        print(f"Error regenerating chart: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        })


# ==================== HTML 生成函數 ====================

def generate_index_html():
    """生成主頁面 HTML"""

    html = '''
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AutoZ Wafer4P Aligner V12.0</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            body {
                font-family: "Microsoft JhengHei", "Segoe UI", Arial, sans-serif;
                background: linear-gradient(135deg, #F5F5F5 0%, #E8E8E8 100%);
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                padding: 20px;
            }

            .container {
                background: #FFFFFF;
                border-radius: 20px;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04);
                max-width: 600px;
                width: 100%;
                padding: 40px;
                animation: fadeIn 0.5s ease-in;
            }

            @keyframes fadeIn {
                from {
                    opacity: 0;
                    transform: translateY(20px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }

            .header {
                text-align: center;
                margin-bottom: 40px;
                padding-bottom: 25px;
                border-bottom: 1px solid #E8E8E8;
            }

            .header h1 {
                font-size: 28px;
                color: #2C2C2C;
                margin-bottom: 10px;
                font-weight: 700;
                letter-spacing: -0.5px;
            }

            .header p {
                color: #666666;
                font-size: 15px;
                font-weight: 400;
            }

            .version-badge {
                display: inline-block;
                background: linear-gradient(135deg, #4A4A4A 0%, #2C2C2C 100%);
                color: white;
                padding: 6px 16px;
                border-radius: 20px;
                font-size: 12px;
                margin-top: 10px;
                font-weight: 600;
                letter-spacing: 0.5px;
            }

            .section {
                margin-bottom: 30px;
            }

            .section-title {
                font-size: 12px;
                color: #666666;
                margin-bottom: 15px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 1px;
            }

            .card {
                background: #FAFAFA;
                border: 1px solid #E8E8E8;
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                transition: all 0.3s ease;
            }

            .card:hover {
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
            }

            .card-title {
                font-size: 14px;
                font-weight: 600;
                color: #2C2C2C;
                margin-bottom: 15px;
            }
            
            .form-group {
                margin-bottom: 20px;
            }

            .form-label {
                display: block;
                font-size: 13px;
                color: #666666;
                margin-bottom: 10px;
                font-weight: 600;
            }

            .form-select {
                width: 100%;
                padding: 14px 16px;
                border: 2px solid #E8E8E8;
                border-radius: 10px;
                font-size: 14px;
                color: #2C2C2C;
                background: white;
                cursor: pointer;
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                font-family: "Microsoft JhengHei", "Segoe UI", Arial, sans-serif;
            }

            .form-select:hover {
                border-color: #D0D0D0;
                background: #FAFAFA;
            }

            .form-select:focus {
                outline: none;
                border-color: #4A4A4A;
                box-shadow: 0 0 0 3px rgba(74, 74, 74, 0.1);
                background: white;
            }

            .btn {
                width: 100%;
                padding: 16px;
                border: none;
                border-radius: 12px;
                font-size: 15px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                letter-spacing: 0.3px;
            }

            .btn-primary {
                background: linear-gradient(135deg, #4A4A4A 0%, #2C2C2C 100%);
                color: #FFFFFF;
                box-shadow: 0 4px 12px rgba(44, 44, 44, 0.15);
            }

            .btn-primary:hover:not(:disabled) {
                transform: translateY(-2px);
                box-shadow: 0 8px 20px rgba(44, 44, 44, 0.25);
                background: linear-gradient(135deg, #5A5A5A 0%, #3C3C3C 100%);
            }

            .btn-secondary {
                background: linear-gradient(135deg, #3A3A3A 0%, #1A1A1A 100%);
                color: #FFFFFF;
                box-shadow: 0 4px 12px rgba(26, 26, 26, 0.15);
            }

            .btn-secondary:hover:not(:disabled) {
                transform: translateY(-2px);
                box-shadow: 0 8px 20px rgba(26, 26, 26, 0.25);
                background: linear-gradient(135deg, #4A4A4A 0%, #2A2A2A 100%);
            }

            .btn:disabled {
                background: #D0D0D0;
                color: #888888;
                cursor: not-allowed;
                transform: none;
                box-shadow: none;
            }
            
            .btn-icon {
                font-size: 18px;
            }

            .file-info {
                background: #FAFAFA;
                border: 1px solid #E8E8E8;
                border-radius: 12px;
                padding: 18px;
                margin-top: 15px;
                display: none;
            }

            .file-info.show {
                display: block;
                animation: slideDown 0.3s ease;
            }

            @keyframes slideDown {
                from {
                    opacity: 0;
                    max-height: 0;
                }
                to {
                    opacity: 1;
                    max-height: 200px;
                }
            }

            .file-info-item {
                display: flex;
                justify-content: space-between;
                padding: 8px 0;
                border-bottom: 1px solid #E8E8E8;
            }

            .file-info-item:last-child {
                border-bottom: none;
            }

            .file-info-label {
                font-weight: 600;
                color: #666666;
                font-size: 13px;
            }

            .file-name {
                color: #2C2C2C;
                font-weight: 600;
                font-size: 13px;
            }

            .progress-container {
                display: none;
                margin-top: 20px;
            }

            .progress-container.show {
                display: block;
                animation: fadeIn 0.3s ease;
            }

            .progress-label {
                font-size: 13px;
                color: #666666;
                margin-bottom: 10px;
                font-weight: 500;
            }

            .progress-bar-wrapper {
                width: 100%;
                height: 32px;
                background: #E8E8E8;
                border-radius: 16px;
                overflow: hidden;
                position: relative;
                box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.06);
            }

            .progress-fill {
                height: 100%;
                background: linear-gradient(90deg, #5A5A5A 0%, #3A3A3A 100%);
                width: 0%;
                transition: width 0.3s ease;
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                font-weight: 600;
                font-size: 13px;
                letter-spacing: 0.3px;
            }

            .loading-spinner {
                display: inline-block;
                width: 20px;
                height: 20px;
                border: 3px solid rgba(255,255,255,0.3);
                border-radius: 50%;
                border-top-color: white;
                animation: spin 1s ease-in-out infinite;
            }

            @keyframes spin {
                to { transform: rotate(360deg); }
            }
            
            .modal {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0, 0, 0, 0.45);
                backdrop-filter: blur(4px);
                z-index: 1000;
                justify-content: center;
                align-items: center;
            }

            .modal.show {
                display: flex;
                animation: fadeIn 0.3s ease;
            }

            .modal-content {
                background: white;
                border-radius: 16px;
                padding: 32px;
                max-width: 400px;
                width: 90%;
                text-align: center;
                animation: slideUp 0.3s ease;
                box-shadow: 0 12px 48px rgba(0, 0, 0, 0.15);
            }

            @keyframes slideUp {
                from {
                    opacity: 0;
                    transform: translateY(50px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }

            .modal-icon {
                font-size: 48px;
                margin-bottom: 20px;
                color: #4A4A4A;
            }

            .modal-icon.error {
                color: #E74C3C;
            }

            .modal-icon.success {
                color: #27AE60;
            }

            .modal-title {
                font-size: 20px;
                font-weight: 700;
                margin-bottom: 12px;
                color: #2C2C2C;
            }

            .modal-message {
                font-size: 14px;
                color: #666666;
                margin-bottom: 24px;
                white-space: pre-line;
                line-height: 1.6;
            }

            .modal-btn {
                padding: 12px 32px;
                border: none;
                border-radius: 10px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                background: linear-gradient(135deg, #4A4A4A 0%, #2C2C2C 100%);
                color: white;
                letter-spacing: 0.3px;
                transition: all 0.3s ease;
            }

            .modal-btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 6px 16px rgba(44, 44, 44, 0.25);
            }
            
            .step-indicator {
                display: flex;
                justify-content: space-between;
                margin-bottom: 35px;
                position: relative;
                padding: 0 10px;
            }

            .step-indicator::before {
                content: '';
                position: absolute;
                top: 20px;
                left: 22%;
                right: 22%;
                height: 3px;
                background: #E8E8E8;
                z-index: 0;
                border-radius: 2px;
            }

            .step {
                flex: 1;
                text-align: center;
                position: relative;
                z-index: 1;
            }

            .step-circle {
                width: 42px;
                height: 42px;
                border-radius: 50%;
                background: #E8E8E8;
                color: #999999;
                display: flex;
                align-items: center;
                justify-content: center;
                margin: 0 auto 10px;
                font-weight: 700;
                font-size: 16px;
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            }

            .step.active .step-circle {
                background: linear-gradient(135deg, #4A4A4A 0%, #2C2C2C 100%);
                color: white;
                box-shadow: 0 4px 12px rgba(44, 44, 44, 0.25);
                transform: scale(1.1);
            }

            .step.completed .step-circle {
                background: linear-gradient(135deg, #27AE60 0%, #229954 100%);
                color: white;
                box-shadow: 0 4px 12px rgba(39, 174, 96, 0.25);
            }

            .step.completed .step-circle::before {
                content: '✓';
                font-size: 20px;
            }

            .step-label {
                font-size: 12px;
                color: #999999;
                font-weight: 500;
                letter-spacing: 0.3px;
            }

            .step.active .step-label {
                color: #2C2C2C;
                font-weight: 700;
            }

            .step.completed .step-label {
                color: #27AE60;
                font-weight: 600;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>AutoZ Wafer4P Aligner</h1>
                <p>Data Visualization Analytics Tool</p>
                <span class="version-badge">Version 12.0</span>
            </div>

            <!-- 步驟指示器 -->
            <div class="step-indicator">
                <div class="step active" id="step1">
                    <div class="step-circle">1</div>
                    <div class="step-label">Machine</div>
                </div>
                <div class="step" id="step2">
                    <div class="step-circle">2</div>
                    <div class="step-label">AutoZLog</div>
                </div>
                <div class="step" id="step3">
                    <div class="step-circle">3</div>
                    <div class="step-label">ALL.txt</div>
                </div>
            </div>

            <!-- Step 1: 機台選擇 -->
            <div class="section">
                <div class="section-title">Step 1: Select Machine Type</div>
                <div class="card">
                    <div class="form-group">
                        <label class="form-label">Machine Type</label>
                        <select class="form-select" id="machineSelect">
                            <option value="">-- Select Machine --</option>
                            <option value="J750">J750</option>
                            <option value="J750EX">J750EX</option>
                            <option value="UFLEX">UFLEX</option>
                            <option value="ETS88">ETS88</option>
                            <option value="Accotest">Accotest</option>
                            <option value="AG93000">AG93000</option>
                            <option value="T2K">T2K</option>
                        </select>
                    </div>
                </div>
            </div>

            <!-- Step 2: AutoZLog.txt -->
            <div class="section">
                <div class="section-title">Step 2: Select AutoZLog.txt</div>
                <button class="btn btn-primary" id="selectAutoZLog" disabled>
                    <i class="fas fa-folder-open"></i>
                    <span>Select AutoZLog.txt</span>
                </button>
                <div class="file-info" id="autoZLogInfo">
                    <div class="file-info-item">
                        <span class="file-info-label">Selected File:</span>
                        <span class="file-name" id="autoZLogFileName">None</span>
                    </div>
                </div>
            </div>

            <!-- Step 3: ALL.txt -->
            <div class="section">
                <div class="section-title">Step 3: Select ALL.txt</div>
                <button class="btn btn-primary" id="selectAllTxt" disabled>
                    <i class="fas fa-folder-open"></i>
                    <span>Select ALL.txt</span>
                </button>
                <div class="file-info" id="allTxtInfo">
                    <div class="file-info-item">
                        <span class="file-info-label">Selected File:</span>
                        <span class="file-name" id="allTxtFileName">None</span>
                    </div>
                </div>
            </div>

            <!-- 進度顯示 -->
            <div class="progress-container" id="progressContainer">
                <div class="progress-label" id="progressLabel">Processing...</div>
                <div class="progress-bar-wrapper">
                    <div class="progress-fill" id="progressFill">0%</div>
                </div>
            </div>
        </div>
        
        <!-- 版本更新 Modal -->
        <div class="modal" id="versionModal">
            <div class="modal-content">
                <div class="modal-icon">
                    <i class="fas fa-info-circle"></i>
                </div>
                <div class="modal-title">Version Update Available</div>
                <div class="modal-message" id="versionMessage"></div>
                <button class="modal-btn" id="versionUpdateBtn">Update Now</button>
            </div>
        </div>

        <!-- 錯誤 Modal -->
        <div class="modal" id="errorModal">
            <div class="modal-content">
                <div class="modal-icon error">
                    <i class="fas fa-exclamation-circle"></i>
                </div>
                <div class="modal-title">Error</div>
                <div class="modal-message" id="errorMessage"></div>
                <button class="modal-btn" onclick="closeModal('errorModal')">OK</button>
            </div>
        </div>

        <!-- 成功 Modal -->
        <div class="modal" id="successModal">
            <div class="modal-content">
                <div class="modal-icon success">
                    <i class="fas fa-check-circle"></i>
                </div>
                <div class="modal-title">Success</div>
                <div class="modal-message" id="successMessage"></div>
                <button class="modal-btn" onclick="closeModal('successModal')">OK</button>
            </div>
        </div>
        
        <script>
            // ========== 版本檢查機制 ==========
            
            async function checkVersionOnStartup() {
                try {
                    const response = await fetch('/api/check_version');
                    const result = await response.json();

                    if (result.status === 'update') {
                        // 顯示更新提示
                        document.getElementById('versionMessage').textContent = result.message;
                        document.getElementById('versionModal').classList.add('show');

                        // 設定更新按鈕事件
                        document.getElementById('versionUpdateBtn').onclick = async () => {
                            document.getElementById('versionModal').classList.remove('show');

                            try {
                                // 呼叫更新 API
                                await fetch('/api/execute_update', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    body: JSON.stringify({ latest_exe: result.latest_exe }),
                                    keepalive: true
                                });
                            } catch (error) {
                                console.error('Failed to execute update:', error);
                            }

                            // 關閉視窗
                            window.close();
                        };
                    } else if (result.status === 'error') {
                        showError(result.message);
                        
                        // 如果是權限或版本錯誤,禁用所有操作
                        if (result.type === 'permission' || result.type === 'version') {
                            document.querySelectorAll('button, select').forEach(el => {
                                el.disabled = true;
                            });
                        }
                    }
                } catch (error) {
                    console.error('Version check failed:', error);
                }
            }

            // 頁面載入時執行版本檢查
            window.addEventListener('DOMContentLoaded', checkVersionOnStartup);
            
            // ========== 主要功能變數 ==========
            
            let selectedMachine = '';
            let autoZLogPath = '';
            let allTxtPath = '';
            
            // 機台選擇
            document.getElementById('machineSelect').addEventListener('change', async function() {
                const machine = this.value;
                
                if (!machine) {
                    return;
                }
                
                try {
                    const response = await fetch('/api/select_machine', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ machine_type: machine })
                    });
                    
                    const result = await response.json();
                    
                    if (result.success) {
                        selectedMachine = machine;
                        document.getElementById('selectAutoZLog').disabled = false;
                        
                        // 更新步驟指示器
                        document.getElementById('step1').classList.add('completed');
                        document.getElementById('step2').classList.add('active');
                    } else {
                        showError(result.error);
                    }
                } catch (error) {
                    showError('Failed to select machine: ' + error);
                }
            });
            
            // 選擇 AutoZLog.txt
            document.getElementById('selectAutoZLog').addEventListener('click', async function() {
                if (this.disabled) return;
                
                try {
                    // 選擇檔案
                    const selectResponse = await fetch('/api/select_file', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_type: 'AutoZLog.txt' })
                    });
                    
                    const selectResult = await selectResponse.json();
                    
                    if (!selectResult.success) {
                        if (selectResult.error !== 'No file selected') {
                            showError(selectResult.error);
                        }
                        return;
                    }
                    
                    autoZLogPath = selectResult.file_path;
                    document.getElementById('autoZLogFileName').textContent = selectResult.file_name;
                    document.getElementById('autoZLogInfo').classList.add('show');
                    
                    // 顯示進度
                    showProgress('Processing AutoZLog.txt...');
                    this.disabled = true;
                    
                    // 處理檔案
                    const processResponse = await fetch('/api/process_autoz_log', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_path: autoZLogPath })
                    });
                    
                    const processResult = await processResponse.json();
                    
                    hideProgress();
                    
                    if (processResult.success) {
                        document.getElementById('selectAllTxt').disabled = false;
                        
                        // 更新步驟指示器
                        document.getElementById('step2').classList.add('completed');
                        document.getElementById('step3').classList.add('active');
                        
                        showSuccess('AutoZLog.txt processed successfully');
                    } else {
                        this.disabled = false;
                        showError(processResult.error);
                    }
                } catch (error) {
                    hideProgress();
                    this.disabled = false;
                    showError('Failed to process AutoZLog.txt: ' + error);
                }
            });
            
            // 選擇 ALL.txt
            document.getElementById('selectAllTxt').addEventListener('click', async function() {
                if (this.disabled) return;
                
                try {
                    // 選擇檔案
                    const selectResponse = await fetch('/api/select_file', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_type: 'ALL.txt' })
                    });
                    
                    const selectResult = await selectResponse.json();
                    
                    if (!selectResult.success) {
                        if (selectResult.error !== 'No file selected') {
                            showError(selectResult.error);
                        }
                        return;
                    }
                    
                    allTxtPath = selectResult.file_path;
                    document.getElementById('allTxtFileName').textContent = selectResult.file_name;
                    document.getElementById('allTxtInfo').classList.add('show');
                    
                    // 顯示進度
                    showProgress('Processing ALL.txt and generating charts...');
                    this.disabled = true;
                    
                    // 處理檔案
                    const processResponse = await fetch('/api/process_all_txt', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_path: allTxtPath })
                    });
                    
                    const processResult = await processResponse.json();
                    
                    hideProgress();
                    
                    if (processResult.success) {
                        // 更新步驟指示器
                        document.getElementById('step3').classList.add('completed');
                        
                        // 跳轉到結果頁面
                        window.location.href = processResult.redirect_url;
                    } else {
                        this.disabled = false;
                        showError(processResult.error);
                    }
                } catch (error) {
                    hideProgress();
                    this.disabled = false;
                    showError('Failed to process ALL.txt: ' + error);
                }
            });
            
            // 顯示/隱藏進度
            function showProgress(message) {
                document.getElementById('progressLabel').textContent = message;
                document.getElementById('progressContainer').classList.add('show');
                const progressFill = document.getElementById('progressFill');
                progressFill.style.width = '0%';
                progressFill.textContent = '0%';

                // 模擬進度動畫
                let progress = 0;
                const interval = setInterval(() => {
                    progress += 2;
                    if (progress > 90) {
                        clearInterval(interval);
                    }
                    progressFill.style.width = progress + '%';
                    progressFill.textContent = progress + '%';
                }, 100);
            }

            function hideProgress() {
                const progressFill = document.getElementById('progressFill');
                progressFill.style.width = '100%';
                progressFill.textContent = '100%';
                setTimeout(() => {
                    document.getElementById('progressContainer').classList.remove('show');
                }, 500);
            }
            
            // Modal 控制
            function showError(message) {
                document.getElementById('errorMessage').textContent = message;
                document.getElementById('errorModal').classList.add('show');
            }
            
            function showSuccess(message) {
                document.getElementById('successMessage').textContent = message;
                document.getElementById('successModal').classList.add('show');
            }
            
            function closeModal(modalId) {
                document.getElementById(modalId).classList.remove('show');
            }
            
            // 點擊 Modal 背景關閉
            window.addEventListener('click', (e) => {
                if (e.target.id === 'errorModal') {
                    document.getElementById('errorModal').classList.remove('show');
                }
                if (e.target.id === 'successModal') {
                    document.getElementById('successModal').classList.remove('show');
                }
                // 版本更新 Modal 不允許點擊背景關閉
            });
            
            // 心跳機制
            setInterval(() => {
                fetch('/api/heartbeat', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' }
                }).catch(error => {
                    console.error('Heartbeat failed:', error);
                });
            }, 60000);
        </script>
    </body>
    </html>
    '''
    
    return html


def generate_result_html(data):
    """生成結果頁面 HTML"""
    
    wafer_data = data['wafer_data']
    x_standard = data['x_standard']
    y_standard = data['y_standard']
    z_standard = data['z_standard']
    
    # 準備 AutoZ complete 點位資料
    standard_point_data = {
        'x': x_standard,
        'y': y_standard,
        'z': z_standard
    }
    
    # 生成所有圖表
    x_fig, x_stats = create_line_chart(wafer_data, 'x', x_standard, standard_point_data)
    y_fig, y_stats = create_line_chart(wafer_data, 'y', y_standard, standard_point_data)
    z_fig, z_stats = create_line_chart(wafer_data, 'z', z_standard, standard_point_data)
    z_anomaly_fig, z_anomaly_stats = create_z_anomaly_chart(wafer_data, z_standard, standard_point_data)
    wafer_status_html = create_wafer_status_dashboard(wafer_data, z_standard)
    
    # 準備圖表資料
    charts_data = {
        'x_chart': x_fig.to_dict(),
        'y_chart': y_fig.to_dict(),
        'z_chart': z_fig.to_dict(),
        'z_anomaly_chart': z_anomaly_fig.to_dict(),
        'x_stats': x_stats,
        'y_stats': y_stats,
        'z_stats': z_stats,
        'z_anomaly_stats': z_anomaly_stats
    }
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    html = f'''
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AutoZ Analysis Results - {selected_machine_type}</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            
            body {{
                font-family: "Microsoft JhengHei", "Segoe UI", Arial, sans-serif;
                background-color: #f5f5f5;
                padding: 20px;
            }}
            
            .container {{
                max-width: 1400px;
                margin: 0 auto;
                background: white;
                border-radius: 12px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                overflow: hidden;
            }}
            
            .header {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 30px;
                text-align: center;
            }}
            
            .header h1 {{
                font-size: 28px;
                margin-bottom: 5px;
            }}
            
            .header p {{
                font-size: 14px;
                opacity: 0.9;
            }}
            
            .timestamp {{
                font-size: 12px;
                text-align: right;
                padding: 10px 30px;
                background: #f8f9fa;
                border-bottom: 1px solid #e0e0e0;
            }}
            
            .tabs {{
                display: flex;
                background: #f8f9fa;
                border-bottom: 2px solid #e0e0e0;
                padding: 0 30px;
            }}
            
            .tab {{
                padding: 15px 30px;
                cursor: pointer;
                font-weight: 600;
                color: #666;
                border-bottom: 3px solid transparent;
                transition: all 0.3s ease;
            }}
            
            .tab:hover {{
                color: #667eea;
            }}
            
            .tab.active {{
                color: #667eea;
                border-bottom-color: #667eea;
                background: white;
            }}
            
            .tab-content {{
                display: none;
                padding: 30px;
            }}
            
            .tab-content.active {{
                display: block;
            }}
            
            .chart-controls {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 8px;
            }}
            
            .chart-title {{
                font-size: 20px;
                font-weight: 600;
                color: #333;
            }}
            
            .axis-selector {{
                display: flex;
                align-items: center;
                gap: 10px;
            }}
            
            .axis-selector label {{
                font-weight: 600;
                color: #666;
            }}
            
            .axis-selector select {{
                padding: 8px 15px;
                border: 2px solid #e0e0e0;
                border-radius: 6px;
                font-size: 14px;
                cursor: pointer;
                background: white;
                min-width: 120px;
            }}
            
            .axis-selector select:focus {{
                outline: none;
                border-color: #667eea;
            }}
            
            .chart-container {{
                margin-bottom: 30px;
                padding: 20px;
                background: white;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            }}
            
            .stats-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 15px;
                margin-bottom: 20px;
            }}
            
            .stat-card {{
                background: #f8f9fa;
                padding: 15px;
                border-radius: 8px;
                text-align: center;
            }}
            
            .stat-label {{
                font-size: 12px;
                color: #666;
                margin-bottom: 5px;
                font-weight: 600;
                text-transform: uppercase;
            }}
            
            .stat-value {{
                font-size: 20px;
                font-weight: 700;
                color: #333;
            }}
            
            .loading {{
                display: none;
                text-align: center;
                padding: 40px;
            }}
            
            .loading.show {{
                display: block;
            }}
            
            .loading-spinner {{
                width: 50px;
                height: 50px;
                border: 5px solid #f3f3f3;
                border-top: 5px solid #667eea;
                border-radius: 50%;
                animation: spin 1s linear infinite;
                margin: 0 auto 15px;
            }}
            
            @keyframes spin {{
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }}
            
            .section-divider {{
                height: 2px;
                background: #e0e0e0;
                margin: 30px 0;
            }}
            
            .footer {{
                text-align: center;
                padding: 20px;
                background: #f8f9fa;
                border-top: 1px solid #e0e0e0;
                color: #666;
                font-size: 12px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>AutoZ Wafer4P Aligner - {selected_machine_type}</h1>
                <p>Data Visualization Analytics Tool</p>
            </div>
            
            <div class="timestamp">
                Generated on: {timestamp}
            </div>
            
            <div class="tabs">
                <div class="tab" onclick="showTab('wafer-status')">Wafer Status</div>
                <div class="tab active" onclick="showTab('chart')">Chart</div>
            </div>
            
            <!-- Wafer Status Tab -->
            <div id="wafer-status" class="tab-content">
                {wafer_status_html}
            </div>
            
            <!-- Chart Tab -->
            <div id="chart" class="tab-content active">
                <!-- X/Y/Z 圖表區域 -->
                <div class="chart-controls">
                    <div class="chart-title">AutoZ Values</div>
                    <div class="axis-selector">
                        <label>Select Axis:</label>
                        <select id="axisSelector">
                            <option value="z" selected>Z Axis</option>
                            <option value="x">X Axis</option>
                            <option value="y">Y Axis</option>
                        </select>
                    </div>
                </div>
                
                <!-- 統計數據 -->
                <div class="stats-container" id="statsContainer">
                    <div class="stat-card">
                        <div class="stat-label">Min</div>
                        <div class="stat-value" id="statMin">{z_stats['min']:.4f} µm</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label">Max</div>
                        <div class="stat-value" id="statMax">{z_stats['max']:.4f} µm</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label">Mean</div>
                        <div class="stat-value" id="statMean">{z_stats['mean']:.4f} µm</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label">Median</div>
                        <div class="stat-value" id="statMedian">{z_stats['median']:.4f} µm</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label">Std Dev</div>
                        <div class="stat-value" id="statStd">{z_stats['std']:.4f} µm</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label">Data Points</div>
                        <div class="stat-value" id="statCount">{z_stats['count']:,}</div>
                    </div>
                </div>
                
                <!-- 圖表容器 -->
                <div class="chart-container">
                    <div id="chart-container"></div>
                </div>
                
                <div class="section-divider"></div>
                
                <!-- Z Anomaly Analysis -->
                <div class="chart-container">
                    <div class="chart-title" style="margin-bottom: 20px;">Z Value Anomaly Analysis</div>
                    <div id="z-anomaly-container"></div>
                </div>
                
                <!-- Loading 動畫 -->
                <div class="loading" id="loading">
                    <div class="loading-spinner"></div>
                    <div>Loading chart...</div>
                </div>
            </div>
            
            <div class="footer">
                AutoZ Wafer4P Aligner | Document automatically generated
            </div>
        </div>
        
        <script>
            // 儲存所有圖表資料
            const initialCharts = {{
                x_chart: {json.dumps(charts_data['x_chart'])},
                y_chart: {json.dumps(charts_data['y_chart'])},
                z_chart: {json.dumps(charts_data['z_chart'])},
                z_anomaly_chart: {json.dumps(charts_data['z_anomaly_chart'])},
                x_stats: {json.dumps(charts_data['x_stats'])},
                y_stats: {json.dumps(charts_data['y_stats'])},
                z_stats: {json.dumps(charts_data['z_stats'])},
                z_anomaly_stats: {json.dumps(charts_data['z_anomaly_stats'])}
            }};
            
            // 初始化圖表 (預設顯示 Z 軸)
            Plotly.newPlot('chart-container', 
                          initialCharts.z_chart.data, 
                          initialCharts.z_chart.layout, 
                          initialCharts.z_chart.config);
            
            Plotly.newPlot('z-anomaly-container',
                          initialCharts.z_anomaly_chart.data,
                          initialCharts.z_anomaly_chart.layout,
                          initialCharts.z_anomaly_chart.config);
            
            // Tab 切換功能
            function showTab(tabName) {{
                // 隱藏所有 tab content
                const tabContents = document.querySelectorAll('.tab-content');
                tabContents.forEach(content => {{
                    content.classList.remove('active');
                }});
                
                // 移除所有 tab active 狀態
                const tabs = document.querySelectorAll('.tab');
                tabs.forEach(tab => {{
                    tab.classList.remove('active');
                }});
                
                // 顯示選定的 tab
                document.getElementById(tabName).classList.add('active');
                
                // 設定對應的 tab 為 active
                event.target.classList.add('active');
            }}
            
            // 軸選擇器事件監聽
            document.getElementById('axisSelector').addEventListener('change', async (e) => {{
                const selectedAxis = e.target.value;
                await updateChart(selectedAxis);
            }});
            
            // 更新圖表函數
            async function updateChart(axisType) {{
                showLoading();
                
                try {{
                    const response = await fetch('/api/regenerate_chart', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ axis_type: axisType }})
                    }});
                    
                    const result = await response.json();
                    
                    if (result.success) {{
                        Plotly.react('chart-container',
                                    result.chart.data,
                                    result.chart.layout,
                                    result.chart.config);
                        
                        updateStats(result.stats);
                    }} else {{
                        alert('Error updating chart: ' + result.error);
                    }}
                }} catch (error) {{
                    alert('Failed to update chart: ' + error);
                }} finally {{
                    hideLoading();
                }}
            }}
            
            // 更新統計數據
            function updateStats(stats) {{
                document.getElementById('statMin').textContent = stats.min.toFixed(4) + ' µm';
                document.getElementById('statMax').textContent = stats.max.toFixed(4) + ' µm';
                document.getElementById('statMean').textContent = stats.mean.toFixed(4) + ' µm';
                document.getElementById('statMedian').textContent = stats.median.toFixed(4) + ' µm';
                document.getElementById('statStd').textContent = stats.std.toFixed(4) + ' µm';
                document.getElementById('statCount').textContent = stats.count.toLocaleString();
            }}
            
            // 顯示/隱藏 Loading
            function showLoading() {{
                document.getElementById('loading').classList.add('show');
                document.getElementById('chart-container').style.opacity = '0.5';
            }}
            
            function hideLoading() {{
                document.getElementById('loading').classList.remove('show');
                document.getElementById('chart-container').style.opacity = '1';
            }}
            
            // 心跳機制
            setInterval(() => {{
                fetch('/api/heartbeat', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }}
                }}).catch(error => {{
                    console.error('Heartbeat failed:', error);
                }});
            }}, 60000);
        </script>
    </body>
    </html>
    '''
    
    return html


# ==================== 主程式進入點 ====================

def main():
    """主程式啟動函數"""
    
    print("=" * 50)
    print("AutoZ Wafer4P Aligner V12.0 - Web App")
    print("=" * 50)
    
    save_log()
    
    version_result = check_version()
    if version_result['status'] == 'error' and version_result.get('type') == 'permission':
        print(f"ERROR: {version_result['message']}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    try:
        port = find_available_port()
        print(f"Starting server on port {port}...")
    except RuntimeError as e:
        print(f"ERROR: {e}")
        input("Press Enter to exit...")
        return
    
    activity_thread = threading.Thread(target=check_activity_thread, daemon=True)
    activity_thread.start()
    print("Activity timeout monitor started (30 minutes)")
    
    url = f'http://localhost:{port}'
    print(f"Opening browser: {url}")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    
    print("Starting Flask server...")
    print("=" * 50)
    
    try:
        app.run(host='localhost', port=port, debug=False, threaded=True)
    except Exception as e:
        print(f"ERROR: Failed to start Flask server: {e}")
        input("Press Enter to exit...")


if __name__ == '__main__':
    main()