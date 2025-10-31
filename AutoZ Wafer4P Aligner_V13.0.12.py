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

# Flask 應用程式初始化
app = Flask(__name__)

# 全域變數 
current_worker = None
last_activity_time = time.time()
ACTIVITY_TIMEOUT = 30 * 60  # 30 分鐘無活動自動關閉

# 分析相關全域變數
selected_machine_type = None
autoz_log_timestamp = None
analysis_file_data = None
processor_module = None


# 工具函數 

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

# 圖表生成函數
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
        showlegend=True,
        xaxis=dict(
            title="Sequential Index",
            title_font=dict(family='Arial Black', size=14),
            tickfont=dict(family='Arial Black', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(
            title=f"{axis_type.upper()} Value (µm)",
            title_font=dict(family='Arial Black', size=14),
            tickfont=dict(family='Arial Black', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        legend=dict(
            x=1.1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0.8)',
            bordercolor='lightgray',
            borderwidth=1,
            font=dict(family='Arial Black', size=12)
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
    """創建 Wafer 狀態儀表板,顯示哪些 Wafer 有低於標準的 Z 值"""
    
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

def create_anomaly_chart(wafer_data, axis_type, standard_value, standard_point_data=None):
    """創建突顯低於標準值的圖表（支援 X/Y/Z 三軸）

    Args:
        wafer_data: 晶圓資料字典
        axis_type: 軸類型 ('x', 'y', 'z')
        standard_value: 標準值
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
        values_key = f'{axis_type}_values'
        if values_key in data and data[values_key]:
            values = data[values_key]
            indices = list(range(current_index, current_index + len(values)))

            for idx, val in zip(indices, values):
                wafer_ids.append(wafer_id)
                if val < standard_value:
                    anomaly_indices.append(idx)
                    anomaly_y.append(val)
                else:
                    normal_indices.append(idx)
                    normal_y.append(val)

            current_index += len(values)
    
    # 檢查第一個點是否為 AutoZ complete 點
    is_first_point_autoz = False
    if standard_point_data and axis_type in standard_point_data:
        autoz_value = standard_point_data[axis_type]
        if len(wafer_ids) > 0:
            if normal_indices and normal_indices[0] == 0:
                first_value = normal_y[0]
            elif anomaly_indices and anomaly_indices[0] == 0:
                first_value = anomaly_y[0]
            else:
                first_value = None

            if first_value is not None and abs(first_value - autoz_value) < 0.001:
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
                text=[f"{'AutoZ Complete' if (idx == 0 and is_first_point_autoz) else f'Wafer ID: {wid}'}<br>{axis_type.upper()} Value: {val:.2f} µm"
                      for idx, wid, val in zip(normal_indices, normal_wafer_ids_list, normal_y)],
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
                text=[f"{'AutoZ Complete' if (idx == 0 and is_first_point_autoz) else f'Wafer ID: {wid}'}<br>{axis_type.upper()} Value: {val:.2f} µm"
                      for idx, wid, val in zip(anomaly_indices, anomaly_wafer_ids_list, anomaly_y)],
                hoverinfo='text'
            )
        )
    
    # 添加標準線
    x_range_start = 0
    x_range_end = current_index
    fig.add_trace(
        go.Scatter(
            x=[x_range_start, x_range_end],
            y=[standard_value, standard_value],
            mode='lines',
            name=f"{axis_type.upper()} Standard ({standard_value} µm)",
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
        showlegend=True,
        xaxis=dict(
            title="Sequential Index",
            title_font=dict(family='Arial Black', size=14),
            tickfont=dict(family='Arial Black', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(
            title=f"{axis_type.upper()} Value (µm)",
            title_font=dict(family='Arial Black', size=14),
            tickfont=dict(family='Arial Black', size=12),
            showgrid=True,
            gridcolor='lightgray'
        ),
        legend=dict(
            x=1.1,
            y=1,
            bgcolor='rgba(255, 255, 255, 0.8)',
            bordercolor='lightgray',
            borderwidth=1,
            font=dict(family='Arial Black', size=12)
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


# Worker 函數
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

# Flask 路由 
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


# API 端點
@app.route('/api/heartbeat', methods=['POST'])
def heartbeat():
    """心跳端點,用於保持活動狀態"""
    try:
        update_activity()
        return jsonify({'success': True, 'status': 'ok'})
    except Exception as e:
        print(f"Error in heartbeat: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Heartbeat failed: {str(e)}'
        })

@app.route('/api/check_version', methods=['GET'])
def api_check_version():
    """版本檢查 API"""
    try:
        update_activity()
        result = check_version()
        return jsonify(result)
    except Exception as e:
        print(f"Error in api_check_version: {str(e)}")
        return jsonify({
            'status': 'error',
            'type': 'exception',
            'message': f'Version check failed: {str(e)}'
        })

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
    try:
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

    except Exception as e:
        print(f"Error in select_machine: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Failed to select machine type: {str(e)}'
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
    try:
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

    except Exception as e:
        print(f"Error in api_process_autoz_log: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Failed to process AutoZLog.txt: {str(e)}'
        })

@app.route('/api/process_all_txt', methods=['POST'])
def api_process_all_txt():
    """處理 ALL.txt API"""
    try:
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

    except Exception as e:
        print(f"Error in api_process_all_txt: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Failed to process ALL.txt: {str(e)}'
        })

@app.route('/api/shutdown', methods=['POST'])
def shutdown():
    """立即關閉伺服器（當瀏覽器關閉時調用）"""
    print("Browser closed, shutting down immediately...")

    def delayed_shutdown():
        time.sleep(0.5)
        os._exit(0)

    threading.Thread(target=delayed_shutdown, daemon=True).start()
    return jsonify({'success': True})

@app.route('/api/regenerate_chart', methods=['POST'])
def regenerate_chart():
    """重新生成圖表 API (用於 X/Y/Z 軸切換)"""
    try:
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
        wafer_data = analysis_file_data['wafer_data']
        x_standard = analysis_file_data['x_standard']
        y_standard = analysis_file_data['y_standard']
        z_standard = analysis_file_data['z_standard']

        standard_point_data = {
            'x': x_standard,
            'y': y_standard,
            'z': z_standard
        }

        # 獲取對應軸的標準值
        standard_map = {
            'x': x_standard,
            'y': y_standard,
            'z': z_standard
        }
        standard_value = standard_map.get(axis_type)

        # 根據軸類型決定主圖表是否傳入標準值（僅 Z 軸顯示標準線）
        main_standard = standard_value if axis_type == 'z' else None

        # 生成主圖表
        main_fig, stats = create_line_chart(
            wafer_data,
            axis_type,
            main_standard,
            standard_point_data
        )

        # 生成異常分析圖表
        anomaly_fig, anomaly_stats = create_anomaly_chart(
            wafer_data,
            axis_type,
            standard_value,
            standard_point_data
        )

        # 將圖表轉為字典格式
        main_chart_dict = main_fig.to_dict()
        anomaly_chart_dict = anomaly_fig.to_dict()

        return jsonify({
            'success': True,
            'main_chart': main_chart_dict,
            'anomaly_chart': anomaly_chart_dict,
            'stats': stats,
            'anomaly_stats': anomaly_stats
        })

    except Exception as e:
        print(f"Error in regenerate_chart: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'Failed to regenerate chart: {str(e)}'
        })

def generate_index_html():
    """Generate main page HTML"""
    
    html = '''
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AutoZ Wafer4P Aligner V12.0</title>
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            body {
                font-family: "Noto Sans TC", "Segoe UI", Arial, sans-serif;
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
                font-family: "Noto Sans TC", "Segoe UI", Arial, sans-serif;
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
                color: #4A4A4A;
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

            .modal-buttons {
                display: flex;
                gap: 10px;
                justify-content: center;
            }

            .modal-btn-secondary {
                background: linear-gradient(135deg, #6c757d 0%, #5a6268 100%);
                color: white;
            }

            .modal-btn-secondary:hover {
                background: linear-gradient(135deg, #5a6268 0%, #545b62 100%);
            }

            /* Step indicator */
            .step-indicator {
                display: flex;
                justify-content: space-between;
                align-items: center;
                position: relative;
                margin-bottom: 35px;
                padding: 0 20px;
            }

            /* Connection line background */
            .step-indicator::before {
                content: '';
                position: absolute;
                top: 24px;
                left: calc(16.666% + 20px);
                right: calc(16.666% + 20px);
                height: 4px;
                background: #E8E8E8;
                z-index: 0;
                border-radius: 2px;
            }

            /* Progress connection line (dynamic) - MODIFIED: Green to Dark Gray */
            .step-indicator::after {
                content: '';
                position: absolute;
                top: 24px;
                left: calc(16.666% + 20px);
                right: 100%;
                height: 4px;
                background: linear-gradient(90deg, #5A5A5A 0%, #3A3A3A 100%);
                z-index: 1;
                border-radius: 2px;
                transition: right 0.6s cubic-bezier(0.4, 0, 0.2, 1);
            }

            /* Progress line animation classes */
            .step-indicator.progress-33::after {
                right: 50%;
            }

            .step-indicator.progress-66::after {
                right: calc(16.666% + 20px);
            }

            .step-indicator.progress-100::after {
                right: calc(16.666% + 20px);
            }

            /* Each step item */
            .step {
                flex: 1;
                display: flex;
                flex-direction: column;
                align-items: center;
                position: relative;
                z-index: 2;
            }

            /* Circle container */
            .step-circle {
                width: 48px;
                height: 48px;
                border-radius: 50%;
                background: #E8E8E8;
                color: #999999;
                display: flex;
                align-items: center;
                justify-content: center;
                font-weight: 700;
                font-size: 18px;
                margin-bottom: 12px;
                position: relative;
                transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            }

            /* Number content */
            .step-circle::before {
                content: attr(data-step);
                position: absolute;
                transition: all 0.3s ease;
            }

            /* Check icon (hidden) */
            .step-circle::after {
                content: '✓';
                position: absolute;
                opacity: 0;
                transform: scale(0);
                transition: all 0.3s ease;
                font-size: 24px;
            }

            /* Active state */
            .step.active .step-circle {
                background: linear-gradient(135deg, #4A4A4A 0%, #2C2C2C 100%);
                color: white;
                transform: scale(1.1);
                box-shadow: 0 4px 16px rgba(44, 44, 44, 0.3);
            }

            /* Completed state - MODIFIED: Green to Dark Gray */
            .step.completed .step-circle {
                background: linear-gradient(135deg, #5A5A5A 0%, #3A3A3A 100%);
                color: white;
                box-shadow: 0 4px 16px rgba(90, 90, 90, 0.3);
            }

            .step.completed .step-circle::before {
                opacity: 0;
                transform: scale(0);
            }

            .step.completed .step-circle::after {
                opacity: 1;
                transform: scale(1);
            }

            /* Step label */
            .step-label {
                font-size: 13px;
                color: #999999;
                font-weight: 500;
                letter-spacing: 0.3px;
                transition: all 0.3s ease;
            }

            .step.active .step-label {
                color: #2C2C2C;
                font-weight: 700;
                font-size: 14px;
            }

            /* MODIFIED: Completed step label color to Dark Gray */
            .step.completed .step-label {
                color: #5A5A5A;
                font-weight: 600;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>AutoZ Wafer4P Aligner</h1>
                <p>Data Visualization Analytics Tool</p>
            </div>

            <!-- Step indicator -->
            <div class="step-indicator" id="stepIndicator">
                <div class="step active" id="step1">
                    <div class="step-circle" data-step="1"></div>
                    <div class="step-label">Machine</div>
                </div>
                <div class="step" id="step2">
                    <div class="step-circle" data-step="2"></div>
                    <div class="step-label">AutoZLog</div>
                </div>
                <div class="step" id="step3">
                    <div class="step-circle" data-step="3"></div>
                    <div class="step-label">ALL.txt</div>
                </div>
            </div>

            <!-- Step 1: Machine selection -->
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

            <!-- Progress display -->
            <div class="progress-container" id="progressContainer">
                <div class="progress-label" id="progressLabel">Processing...</div>
                <div class="progress-bar-wrapper">
                    <div class="progress-fill" id="progressFill">0%</div>
                </div>
            </div>
        </div>
        
        <!-- Version update modal -->
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

        <!-- Error modal -->
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

        <!-- Success modal -->
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

        <!-- Message modal -->
        <div class="modal" id="messageModal">
            <div class="modal-content">
                <div class="modal-icon" id="messageIcon">
                    <i class="fa-solid fa-circle-info"></i>
                </div>
                <div class="modal-title" id="messageTitle">Message</div>
                <div class="modal-message" id="messageText">Message content</div>
                <div class="modal-buttons">
                    <button class="modal-btn" id="messageOkBtn">OK</button>
                </div>
            </div>
        </div>

        <!-- Confirm modal -->
        <div class="modal" id="confirmModal">
            <div class="modal-content">
                <div class="modal-icon">
                    <i class="fas fa-question-circle"></i>
                </div>
                <div class="modal-title">Confirm</div>
                <div class="modal-message" id="confirmMessage">確認訊息</div>
                <div class="modal-buttons">
                    <button class="modal-btn modal-btn-secondary" id="confirmCancelBtn">Cancel</button>
                    <button class="modal-btn" id="confirmOkBtn">Confirm</button>
                </div>
            </div>
        </div>

        <script>
            // ========== Browser close detection mechanism ==========

            let isNormalNavigation = false;

            window.addEventListener('beforeunload', () => {
                if (!isNormalNavigation) {
                    fetch('/api/shutdown', {
                        method: 'POST',
                        keepalive: true
                    });
                }
            });

            // ========== Promise-based dialog functions ==========

            async function showConfirm(message, isAlert = false) {
                return new Promise((resolve) => {
                    const confirmModal = document.getElementById('confirmModal');
                    const confirmMessage = document.getElementById('confirmMessage');
                    const confirmOkBtn = document.getElementById('confirmOkBtn');
                    const confirmCancelBtn = document.getElementById('confirmCancelBtn');

                    confirmMessage.textContent = message;

                    if (isAlert) {
                        confirmCancelBtn.style.display = 'none';
                        confirmOkBtn.textContent = 'OK';
                    } else {
                        confirmCancelBtn.style.display = 'block';
                        confirmOkBtn.textContent = 'Confirm';
                    }

                    confirmModal.classList.add('show');

                    confirmOkBtn.onclick = () => {
                        confirmModal.classList.remove('show');
                        resolve(true);
                    };

                    confirmCancelBtn.onclick = () => {
                        confirmModal.classList.remove('show');
                        resolve(false);
                    };
                });
            }

            function showMessage(title, message, iconClass = 'fa-circle-info', iconColor = '#4A4A4A') {
                const messageModal = document.getElementById('messageModal');
                const messageIcon = document.getElementById('messageIcon');
                const messageTitle = document.getElementById('messageTitle');
                const messageText = document.getElementById('messageText');
                const messageOkBtn = document.getElementById('messageOkBtn');

                messageIcon.innerHTML = `<i class="fa-solid ${iconClass}"></i>`;
                messageIcon.style.color = iconColor;
                messageTitle.textContent = title;
                messageText.textContent = message;

                messageModal.classList.add('show');

                messageOkBtn.onclick = () => {
                    messageModal.classList.remove('show');
                };
            }

            // ========== Version check mechanism ==========

            async function checkVersionOnStartup() {
                try {
                    const response = await fetch('/api/check_version');
                    const result = await response.json();

                    if (result.status === 'update') {
                        document.getElementById('versionMessage').textContent = result.message;
                        document.getElementById('versionModal').classList.add('show');

                        document.getElementById('versionUpdateBtn').onclick = async () => {
                            document.getElementById('versionModal').classList.remove('show');

                            try {
                                await fetch('/api/execute_update', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    body: JSON.stringify({ latest_exe: result.latest_exe }),
                                    keepalive: true
                                });
                            } catch (error) {
                                console.error('Failed to execute update:', error);
                            }

                            window.close();
                        };
                    } else if (result.status === 'error') {
                        showError(result.message);
                        
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

            window.addEventListener('DOMContentLoaded', checkVersionOnStartup);
            
            // ========== Main variables ==========
            
            let selectedMachine = '';
            let autoZLogPath = '';
            let allTxtPath = '';
            let progressInterval = null;  // ADDED: Global progress interval variable
            
            // Machine selection
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
                        
                        document.getElementById('step1').classList.add('completed');
                        document.getElementById('step2').classList.add('active');
                        document.getElementById('stepIndicator').classList.add('progress-33');
                    } else {
                        showError(result.error);
                    }
                } catch (error) {
                    showError('Failed to select machine: ' + error);
                }
            });
            
            // Select AutoZLog.txt
            document.getElementById('selectAutoZLog').addEventListener('click', async function() {
                if (this.disabled) return;
                
                try {
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
                    
                    showProgress('Processing AutoZLog.txt...');
                    this.disabled = true;
                    
                    const processResponse = await fetch('/api/process_autoz_log', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_path: autoZLogPath })
                    });
                    
                    const processResult = await processResponse.json();
                    
                    hideProgress();
                    
                    if (processResult.success) {
                        document.getElementById('selectAllTxt').disabled = false;
                        
                        document.getElementById('step2').classList.add('completed');
                        document.getElementById('step3').classList.add('active');
                        document.getElementById('stepIndicator').classList.remove('progress-33');
                        document.getElementById('stepIndicator').classList.add('progress-66');
                        
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
            
            // Select ALL.txt
            document.getElementById('selectAllTxt').addEventListener('click', async function() {
                if (this.disabled) return;
                
                try {
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
                    
                    showProgress('Processing ALL.txt and generating charts...');
                    this.disabled = true;
                    
                    const processResponse = await fetch('/api/process_all_txt', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ file_path: allTxtPath })
                    });
                    
                    const processResult = await processResponse.json();
                    
                    hideProgress();
                    
                    if (processResult.success) {
                        document.getElementById('step3').classList.add('completed');
                        document.getElementById('stepIndicator').classList.remove('progress-66');
                        document.getElementById('stepIndicator').classList.add('progress-100');

                        isNormalNavigation = true;  // 標記為正常跳轉
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
            
            // MODIFIED: Fixed showProgress to prevent duplicate progress bar execution
            function showProgress(message) {
                // Clear old progress interval to prevent duplicate execution
                if (progressInterval) {
                    clearInterval(progressInterval);
                    progressInterval = null;
                }

                document.getElementById('progressLabel').textContent = message;
                document.getElementById('progressContainer').classList.add('show');
                const progressFill = document.getElementById('progressFill');
                progressFill.style.width = '0%';
                progressFill.textContent = '0%';

                // Simulate progress animation
                let progress = 0;
                progressInterval = setInterval(() => {
                    progress += 2;
                    if (progress > 90) {
                        clearInterval(progressInterval);
                        progressInterval = null;
                    }
                    progressFill.style.width = progress + '%';
                    progressFill.textContent = progress + '%';
                }, 100);
            }

            // MODIFIED: Fixed hideProgress to ensure interval is cleared
            function hideProgress() {
                // Clear progress interval
                if (progressInterval) {
                    clearInterval(progressInterval);
                    progressInterval = null;
                }
                
                const progressFill = document.getElementById('progressFill');
                progressFill.style.width = '100%';
                progressFill.textContent = '100%';
                setTimeout(() => {
                    document.getElementById('progressContainer').classList.remove('show');
                }, 500);
            }
            
            // Modal controls
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
            
            // Click modal background to close
            window.addEventListener('click', (e) => {
                if (e.target.id === 'errorModal') {
                    document.getElementById('errorModal').classList.remove('show');
                }
                if (e.target.id === 'successModal') {
                    document.getElementById('successModal').classList.remove('show');
                }
                if (e.target.id === 'messageModal') {
                    document.getElementById('messageModal').classList.remove('show');
                }
                if (e.target.id === 'confirmModal') {
                    document.getElementById('confirmModal').classList.remove('show');
                }
            });
            
            // Heartbeat mechanism
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
    """生成結果頁面 HTML - 採用現代化設計風格"""

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

    # 為所有三個軸生成圖表(包含 AutoZ complete 點位)
    x_fig, x_stats = create_line_chart(wafer_data, 'x', x_standard, standard_point_data)
    y_fig, y_stats = create_line_chart(wafer_data, 'y', y_standard, standard_point_data)
    z_fig, z_stats = create_line_chart(wafer_data, 'z', z_standard, standard_point_data)

    # 生成新的圖表(包含 AutoZ complete 點位)
    z_anomaly_fig, z_anomaly_stats = create_anomaly_chart(wafer_data, 'z', z_standard, standard_point_data)
    wafer_status_html = create_wafer_status_dashboard(wafer_data, z_standard)

    # 轉換為 HTML
    x_html = x_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    y_html = y_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    z_html = z_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})
    z_anomaly_html = z_anomaly_fig.to_html(include_plotlyjs=False, full_html=False, config={"responsive": True})

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 計算統計資訊
    total_wafers = len(wafer_data)
    total_points = sum(len(data.get('z_values', [])) for data in wafer_data.values())
    anomaly_count = z_anomaly_stats.get('anomaly_points', 0)
    anomaly_percent = z_anomaly_stats.get('anomaly_percent', 0)

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
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AutoZ Wafer4P Aligner - {selected_machine_type}</title>
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>

            body {{
                font-family: "Noto Sans TC", Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #f8f9fa;
            }}

            .container {{
                width: 98%;
                margin: 20px auto;
            }}

            /* Header Styling */
            .header {{
                background-color: #2D2D2D;
                color: #E0E0E0;
                padding: 15px;
                text-align: center;
                border-radius: 8px 8px 0 0;
                margin-bottom: -10px;
            }}

            /* Tab System */
            .tabs {{
                display: flex;
                background-color: #2D2D2D;
                padding: 10px 10px 0 10px;
                border-radius: 8px 8px 0 0;
            }}

            .tab {{
                padding: 10px 20px;
                cursor: pointer;
                background-color: #454545;
                color: #E0E0E0;
                margin-right: 5px;
                border-radius: 8px 8px 0 0;
                font-weight: bold;
                transition: all 0.3s ease;
            }}

            .tab:hover {{
                background-color: #5A5A5A;
            }}

            .tab.active {{
                background-color: #f8f9fa;
                color: #333;
            }}

            .tab-content {{
                display: none;
                padding: 20px;
                background-color: white;
                border-radius: 0 0 8px 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            }}

            .tab-content.active {{
                display: block;
            }}

            /* Info Section */
            .info-section {{
                margin-bottom: 30px;
                padding: 20px;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                width: 95%;
                max-width: 1100px;
                margin-left: auto;
                margin-right: auto;
            }}

            .info-title {{
                font-size: 20px;
                font-weight: bold;
                margin-bottom: 20px;
                color: #333;
                border-bottom: 2px solid #e0e0e0;
                padding-bottom: 10px;
            }}

            .info-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
                gap: 15px;
                margin-bottom: 20px;
            }}

            .info-item {{
                padding: 15px;
                background-color: #f8f9fa;
                border-radius: 8px;
                border-left: 4px solid #4A4A4A;
            }}

            .info-label {{
                font-weight: 600;
                color: #666;
                margin-bottom: 8px;
                font-size: 13px;
            }}

            .info-value {{
                color: #333;
                font-size: 18px;
                font-weight: 500;
            }}

            .summary-box {{
                background-color: #D4EDDA;
                border-left: 4px solid #28A745;
                padding: 20px;
                margin: 20px 0;
                border-radius: 8px;
            }}

            .summary-box p {{
                margin: 8px 0;
                font-size: 15px;
            }}

            /* Chart Container */
            .chart-container {{
                margin-bottom: 30px;
                padding: 20px;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                width: 98%;
                margin-left: auto;
                margin-right: auto;
            }}

            .chart-title {{
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 15px;
                color: #333;
                text-align: center;
            }}

            .js-plotly-plot, .plot-container {{
                width: 100% !important;
                max-width: 1100px !important;
                margin: 0 auto !important;
            }}

            /* Axis Control Panel */
            .axis-control-panel {{
                margin: 20px auto 30px auto;
                width: 98%;
                background: white;
                border-radius: 12px;
                border: 2px solid #e9ecef;
                overflow: hidden;
                box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            }}

            .control-panel-header {{
                background: linear-gradient(135deg, #2D2D2D 0%, #4A4A4A 100%);
                color: white;
                padding: 18px 24px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                flex-wrap: wrap;
                gap: 15px;
            }}

            .header-left {{
                display: flex;
                align-items: center;
                gap: 12px;
            }}

            .header-icon {{
                font-size: 24px;
            }}

            .header-title {{
                font-size: 18px;
                font-weight: 600;
                letter-spacing: 0.3px;
            }}

            .header-subtitle {{
                font-size: 13px;
                opacity: 0.85;
                margin-top: 4px;
            }}

            .axis-buttons {{
                display: flex;
                gap: 12px;
                flex-wrap: wrap;
            }}

            .axis-btn {{
                padding: 10px 24px;
                border: 2px solid rgba(255,255,255,0.3);
                background: rgba(255,255,255,0.15);
                border-radius: 8px;
                cursor: pointer;
                font-size: 15px;
                font-weight: 600;
                color: white;
                transition: all 0.3s ease;
                font-family: "Noto Sans TC", Arial, sans-serif;
            }}

            .axis-btn:hover {{
                background: rgba(255,255,255,0.25);
                border-color: rgba(255,255,255,0.5);
                transform: translateY(-1px);
            }}

            .axis-btn.active {{
                background: linear-gradient(135deg, #E0E0E0 0%, #FFFFFF 100%);
                color: #2D2D2D;
                border-color: #FFFFFF;
                box-shadow: 0 2px 8px rgba(255, 255, 255, 0.3);
            }}

            /* Statistics Container */
            .statistics-container {{
                margin: 20px auto;
                padding: 20px;
                background-color: #f8f9fa;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                width: 98%;
            }}

            .stats-title {{
                text-align: center;
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 20px;
                color: #333;
            }}

            .statistics-box {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                gap: 15px;
                padding: 10px;
            }}

            .stat-item {{
                padding: 15px;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                text-align: center;
            }}

            .stat-label {{
                font-weight: 600;
                display: block;
                margin-bottom: 8px;
                color: #666;
                font-size: 13px;
            }}

            .stat-value {{
                font-size: 20px;
                color: #333;
                font-weight: 700;
            }}

            /* Section Divider */
            .section-divider {{
                border-top: 2px solid #e9ecef;
                margin: 40px 0;
                width: 100%;
            }}

            /* Loading Overlay */
            .loading-overlay {{
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: rgba(0, 0, 0, 0.6);
                backdrop-filter: blur(4px);
                z-index: 9999;
                justify-content: center;
                align-items: center;
            }}

            .loading-overlay.active {{
                display: flex;
            }}

            .loading-content {{
                background: white;
                padding: 40px;
                border-radius: 16px;
                text-align: center;
                box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            }}

            .spinner {{
                border: 4px solid #f3f3f3;
                border-top: 4px solid #2D2D2D;
                border-radius: 50%;
                width: 50px;
                height: 50px;
                animation: spin 1s linear infinite;
                margin: 0 auto 20px;
            }}

            @keyframes spin {{
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }}

            .loading-text {{
                font-size: 16px;
                font-weight: 600;
                color: #2D2D2D;
                margin-bottom: 8px;
            }}

            .loading-subtext {{
                font-size: 14px;
                color: #6c757d;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <!-- Header -->
            <div class="header">
                <h1>AutoZ Wafer4P Aligner - {selected_machine_type}</h1>
                <p style="font-size: 14px;">Data Visualization Analytics Tool</p>
                <p style="font-size: 14px;">Analysis Time: {timestamp}</p>
            </div>

            <!-- Tab Navigation -->
            <div class="tabs">
                <div class="tab active" onclick="showTab('info')">Info</div>
                <div class="tab" onclick="showTab('wafer-status')">Wafer Status</div>
                <div class="tab" onclick="showTab('charts')">Charts</div>
            </div>

            <!-- Info Tab Content -->
            <div id="info" class="tab-content active">
                <div class="info-section">
                    <div class="info-title">Basic Information</div>
                    <div class="info-grid">
                        <div class="info-item">
                            <div class="info-label">Machine Type</div>
                            <div class="info-value">{selected_machine_type}</div>
                        </div>
                        <div class="info-item">
                            <div class="info-label">Analysis Time</div>
                            <div class="info-value">{timestamp}</div>
                        </div>
                        <div class="info-item">
                            <div class="info-label">Status</div>
                            <div class="info-value">Completed</div>
                        </div>
                    </div>
                </div>

                <div class="info-section">
                    <div class="info-title">Standard Values</div>
                    <div class="info-grid">
                        <div class="info-item">
                            <div class="info-label">X Standard</div>
                            <div class="info-value">{x_standard:.4f} µm</div>
                        </div>
                        <div class="info-item">
                            <div class="info-label">Y Standard</div>
                            <div class="info-value">{y_standard:.4f} µm</div>
                        </div>
                        <div class="info-item">
                            <div class="info-label">Z Standard</div>
                            <div class="info-value">{z_standard:.4f} µm</div>
                        </div>
                    </div>
                </div>

                <div class="info-section">
                    <div class="info-title">Analysis Summary</div>
                    <div class="summary-box">
                        <p><strong>Total Wafers:</strong> {total_wafers}</p>
                        <p><strong>Total Data Points:</strong> {total_points:,}</p>
                        <p><strong>Z Axis Anomalies:</strong> {anomaly_count} ({anomaly_percent:.2f}%)</p>
                        <p><strong>AutoZ Complete Points:</strong> Identified</p>
                    </div>
                </div>
            </div>

            <!-- Wafer Status Tab Content -->
            <div id="wafer-status" class="tab-content">
                {wafer_status_html}
            </div>

            <!-- Charts Tab Content -->
            <div id="charts" class="tab-content">
                <!-- Axis Control Panel -->
                <div class="axis-control-panel">
                    <div class="control-panel-header">
                        <div class="header-left">
                            <span class="header-icon">⚙️</span>
                            <div>
                                <div class="header-title">Axis Selection</div>
                                <div class="header-subtitle">Select axis to display chart</div>
                            </div>
                        </div>
                        <div class="axis-buttons">
                            <button class="axis-btn active" data-axis="z" onclick="switchAxis('z')">Z Axis</button>
                            <button class="axis-btn" data-axis="x" onclick="switchAxis('x')">X Axis</button>
                            <button class="axis-btn" data-axis="y" onclick="switchAxis('y')">Y Axis</button>
                        </div>
                    </div>
                </div>

                <!-- Statistics Section -->
                <div class="statistics-container">
                    <div class="stats-title" id="statsTitle">Z Data Statistics</div>
                    <div id="statsContent">
                        {z_stats_html}
                    </div>
                </div>

                <!-- AutoZ Values Chart -->
                <div class="chart-container" id="autoZValuesChartContainer">
                    <div class="chart-title">Z AutoZ Values</div>
                    {z_html}
                </div>

                <div class="section-divider"></div>

                <!-- Z Value Anomaly Analysis Chart -->
                <div class="chart-container" id="anomalyChartContainer">
                    <div class="chart-title" id="anomalyChartTitle">Z Value Anomaly Analysis</div>
                    <div id="anomalyChart">{z_anomaly_html}</div>
                </div>
            </div>
        </div>

        <!-- Loading Overlay -->
        <div class="loading-overlay" id="loadingOverlay">
            <div class="loading-content">
                <div class="spinner"></div>
                <div class="loading-text">Loading Chart...</div>
                <div class="loading-subtext">Please wait while we update the visualization</div>
            </div>
        </div>

        <script>
            // ========== Browser close detection mechanism ==========

            let isNormalNavigation = false;

            window.addEventListener('beforeunload', () => {{
                if (!isNormalNavigation) {{
                    fetch('/api/shutdown', {{
                        method: 'POST',
                        keepalive: true
                    }});
                }}
            }});

            // Tab switching function
            function showTab(tabName) {{
                // Hide all tab contents
                var tabContents = document.getElementsByClassName("tab-content");
                for (var i = 0; i < tabContents.length; i++) {{
                    tabContents[i].style.display = "none";
                    tabContents[i].classList.remove("active");
                }}

                // Remove active class from all tabs
                var tabs = document.getElementsByClassName("tab");
                for (var i = 0; i < tabs.length; i++) {{
                    tabs[i].classList.remove("active");
                }}

                // Show selected tab and mark button as active
                document.getElementById(tabName).style.display = "block";
                document.getElementById(tabName).classList.add("active");
                event.target.classList.add("active");
            }}

            // Axis switching function
            async function switchAxis(axisType) {{
                // Update button states
                document.querySelectorAll('.axis-btn').forEach(btn => {{
                    btn.classList.remove('active');
                }});
                const clickedBtn = document.querySelector(`button[data-axis="${{axisType}}"]`);
                if (clickedBtn) {{
                    clickedBtn.classList.add('active');
                }}

                // Show loading
                const loadingOverlay = document.getElementById('loadingOverlay');
                loadingOverlay.classList.add('active');

                try {{
                    // Call API to regenerate chart
                    const response = await fetch('/api/regenerate_chart', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ axis_type: axisType }})
                    }});

                    const result = await response.json();

                    if (result.success) {{
                        // Update stats title
                        const statsTitle = document.getElementById('statsTitle');
                        statsTitle.textContent = axisType.toUpperCase() + ' Data Statistics';

                        // Update statistics data
                        const stats = result.stats;
                        const statsHtml = `
                            <div class="statistics-box">
                                <div class="stat-item"><span class="stat-label">${{axisType.toUpperCase()}} Min:</span> <span class="stat-value">${{stats.min.toFixed(4)}} µm</span></div>
                                <div class="stat-item"><span class="stat-label">${{axisType.toUpperCase()}} Max:</span> <span class="stat-value">${{stats.max.toFixed(4)}} µm</span></div>
                                <div class="stat-item"><span class="stat-label">${{axisType.toUpperCase()}} Mean:</span> <span class="stat-value">${{stats.mean.toFixed(4)}} µm</span></div>
                                <div class="stat-item"><span class="stat-label">${{axisType.toUpperCase()}} Median:</span> <span class="stat-value">${{stats.median.toFixed(4)}} µm</span></div>
                                <div class="stat-item"><span class="stat-label">${{axisType.toUpperCase()}} Std Dev:</span> <span class="stat-value">${{stats.std.toFixed(4)}} µm</span></div>
                                <div class="stat-item"><span class="stat-label">Data Points:</span> <span class="stat-value">${{stats.count.toLocaleString()}}</span></div>
                            </div>
                        `;
                        document.getElementById('statsContent').innerHTML = statsHtml;

                        // Update main chart
                        const chartContainer = document.getElementById('autoZValuesChartContainer');
                        chartContainer.innerHTML = '<div class="chart-title">' + axisType.toUpperCase() + ' AutoZ Values</div><div id="newChart"></div>';
                        Plotly.newPlot('newChart', result.main_chart.data, result.main_chart.layout, {{responsive: true}});

                        // Update anomaly chart
                        const anomalyContainer = document.getElementById('anomalyChartContainer');
                        anomalyContainer.innerHTML = `
                            <div class="chart-title" id="anomalyChartTitle">${{axisType.toUpperCase()}} Value Anomaly Analysis</div>
                            <div id="anomalyChart"></div>
                        `;
                        Plotly.newPlot('anomalyChart', result.anomaly_chart.data, result.anomaly_chart.layout, {{responsive: true}});
                    }} else {{
                        console.error('Failed to regenerate chart:', result.error);
                        alert('Failed to load chart: ' + result.error);
                    }}
                }} catch (error) {{
                    console.error('Error switching axis:', error);
                    alert('Error loading chart. Please try again.');
                }} finally {{
                    // Hide loading
                    loadingOverlay.classList.remove('active');
                }}
            }}

            // Heartbeat mechanism
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

def main():
    """主程式啟動函數"""
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