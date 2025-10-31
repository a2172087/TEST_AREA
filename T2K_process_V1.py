import re
from datetime import datetime

def convert_timestamp_format(timestamp):
    """將 '2025/04/08 14:59:41.200' 轉換為 '04.08 14:59:41.200'"""
    try:
        dt = datetime.strptime(timestamp, "%Y/%m/%d %H:%M:%S.%f")
        return f"{dt.month:02d}.{dt.day:02d} {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}.{int(dt.microsecond/1000):03d}"
    except Exception as e:
        print(f"Error converting timestamp: {e}")
        return None

def process_autoz_log(file_path):
    """處理 T2K 的 AutoZLog.txt 檔案以提取所需數據"""
    # 存儲時間戳的變數
    autoz_complete_timestamp = None
    
    # T2K 的搜尋關鍵字
    search_pattern = "AutoZ Completed Successful."
    
    # 讀取檔案
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # 從末尾開始向前搜索 "AutoZ Completed Successful."
    for i in range(len(lines)-1, -1, -1):
        if search_pattern in lines[i]:
            # 提取時間戳（格式：2025-05-12 16:03:01.689）
            timestamp_match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3})', lines[i])
            if timestamp_match:
                timestamp_part = timestamp_match.group(1)
                # 將時間戳格式從 YYYY-MM-DD HH:MM:SS.fff 轉換為 YYYY/MM/DD HH:MM:SS.fff
                timestamp_part = timestamp_part.replace("-", "/")
                autoz_complete_timestamp = timestamp_part
                break
    
    # 只打印最終結果
    if autoz_complete_timestamp is not None:
        print(f"Found AutoZ Completed Successful timestamp: {autoz_complete_timestamp}")
    else:
        print(f"Could not find '{search_pattern}' in log file.")
    
    return autoz_complete_timestamp

def process_all_txt(file_path, autoz_log_timestamp):
    """處理 ALL.TXT 檔案以提取所需數據，只處理 Auto Z complete 點之後的數據"""
    print(f"Processing ALL.TXT file: {file_path}")
    
    # 檢查 AutoZ Complete 時間戳是否可用
    if not autoz_log_timestamp:
        print("AutoZ Complete timestamp not available. Please process AutoZLog.txt first.")
        raise ValueError("AutoZ Complete timestamp not available")
    
    # 將時間戳從 '2025/04/08 14:59:41.200' 轉換為 '04.08 14:59:41.200'
    autoz_complete_converted = convert_timestamp_format(autoz_log_timestamp)
    if not autoz_complete_converted:
        print("Could not convert timestamp format.")
        raise ValueError("Timestamp format conversion failed")
    
    print(f"Searching for Auto Z complete point after {autoz_complete_converted}")
    
    # 讀取檔案
    with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
        lines = file.readlines()
    
    # 找出所有晶圓 ID 及其開始時間
    wafer_ids = {}
    wafer_id_pattern = re.compile(r'(\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\.\d{3}) EX_G S : b([A-Za-z0-9]+[.-][A-Za-z0-9]+)')
    
    for line in lines:
        match = wafer_id_pattern.search(line)
        if match:
            timestamp = match.group(1)
            wafer_id = match.group(2)
            
            # 如果晶圓 ID 已存在，僅在此時間戳較早時更新
            if wafer_id in wafer_ids:
                if timestamp < wafer_ids[wafer_id]['start_time']:
                    wafer_ids[wafer_id]['start_time'] = timestamp
            else:
                wafer_ids[wafer_id] = {
                    'start_time': timestamp,
                    'x_values': [],
                    'y_values': [],
                    'z_values': []
                }
    
    # 按開始時間排序晶圓 ID
    sorted_wafer_ids = sorted(wafer_ids.items(), key=lambda x: x[1]['start_time'])
    
    print(f"Found {len(sorted_wafer_ids)} wafer IDs:")
    for wafer_id, data in sorted_wafer_ids:
        print(f"  - {wafer_id}: {data['start_time']}")
    
    # Auto Z complete 點的值（這些值會作為標準值返回給主程式）
    autoz_complete_x = None
    autoz_complete_y = None
    autoz_complete_z = None
    autoz_complete_timestamp = None
    
    # 尋找 Auto Z complete 點（AutoZ Complete Successful 之後的第一個 ndlp-correct）
    found_autoz_complete = False
    ndlp_pattern = re.compile(r'(\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\.\d{3}).*ndlp-correct: \(\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+) \)')
    
    # 第一次遍歷：找到 Auto Z complete 點
    for line in lines:
        match = ndlp_pattern.search(line)
        if match:
            timestamp = match.group(1)
            
            # AutoZ Complete Successful 之後的第一個 ndlp-correct 就是 Auto Z complete 點
            if timestamp >= autoz_complete_converted and not found_autoz_complete:
                autoz_complete_timestamp = timestamp
                autoz_complete_x = float(match.group(2))
                autoz_complete_y = float(match.group(3))
                autoz_complete_z = float(match.group(4))
                found_autoz_complete = True
                print(f"Found Auto Z complete point at {autoz_complete_timestamp}:")
                print(f"  X={autoz_complete_x}, Y={autoz_complete_y}, Z={autoz_complete_z}")
                print(f"All ndlp-correct data before this point will be ignored.")
                break
    
    # 確保找到了 Auto Z complete 點
    if not found_autoz_complete:
        print("Could not find Auto Z complete point (first ndlp-correct after AutoZ Complete Successful).")
        raise ValueError("Auto Z complete point not found")
    
    # 第二次遍歷：只處理 Auto Z complete 點之後的數據
    print(f"Processing ndlp-correct data after Auto Z complete point...")
    data_count = 0
    
    for line in lines:
        match = ndlp_pattern.search(line)
        if match:
            timestamp = match.group(1)
            
            # 只處理 Auto Z complete 時間點之後的數據（不包含 Auto Z complete 點本身）
            if timestamp >= autoz_complete_timestamp:
                x_val = float(match.group(2))
                y_val = float(match.group(3))
                z_val = float(match.group(4))
                data_count += 1
                
                assigned = False
                for i in range(len(sorted_wafer_ids) - 1):
                    current_wafer_id, current_data = sorted_wafer_ids[i]
                    next_wafer_id, next_data = sorted_wafer_ids[i + 1]
                    
                    if current_data['start_time'] <= timestamp < next_data['start_time']:
                        wafer_ids[current_wafer_id]['x_values'].append(x_val)
                        wafer_ids[current_wafer_id]['y_values'].append(y_val)
                        wafer_ids[current_wafer_id]['z_values'].append(z_val)
                        assigned = True
                        break
                
                # 如果未分配，檢查是否屬於最後一個晶圓
                if not assigned and sorted_wafer_ids:
                    last_wafer_id, last_data = sorted_wafer_ids[-1]
                    if timestamp >= last_data['start_time']:
                        wafer_ids[last_wafer_id]['x_values'].append(x_val)
                        wafer_ids[last_wafer_id]['y_values'].append(y_val)
                        wafer_ids[last_wafer_id]['z_values'].append(z_val)
    
    print(f"Processed {data_count} ndlp-correct entries after Auto Z complete point.")
    
    # 打印提取數據的摘要
    print("\nData extraction summary:")
    for wafer_id, data in wafer_ids.items():
        print(f"Wafer ID: {wafer_id}")
        print(f"  Start time: {data['start_time']}")
        print(f"  Number of ndlp-correct entries: {len(data['x_values'])}")
        if data['x_values']:
            print(f"  X value range: {min(data['x_values']):.3f} to {max(data['x_values']):.3f}")
            print(f"  Y value range: {min(data['y_values']):.3f} to {max(data['y_values']):.3f}")
            print(f"  Z value range: {min(data['z_values']):.3f} to {max(data['z_values']):.3f}")
        print("---")
    
    # 返回處理結果（使用主程式期望的 key 名稱）
    return {
        'wafer_data': wafer_ids,
        'x_standard': autoz_complete_x,      # Auto Z complete 的 X 值作為標準值
        'y_standard': autoz_complete_y,      # Auto Z complete 的 Y 值作為標準值
        'z_standard': autoz_complete_z,      # Auto Z complete 的 Z 值作為標準值
        'standard_timestamp': autoz_complete_timestamp  # Auto Z complete 的時間戳
    }