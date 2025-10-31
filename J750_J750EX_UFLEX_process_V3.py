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
    """處理 AutoZLog.txt 檔案以提取所需數據"""
    # 存儲時間戳的變數
    AutoZLog_Last_Trigger4pinalignment_Time = None
    
    # 讀取檔案
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # 從末尾開始向前搜索 [DisplaySite] 包含 "Last Contact"
    display_site_index = None
    for i in range(len(lines)-1, -1, -1):
        if "[DisplaySite]" in lines[i] and "Last Contact" in lines[i]:
            display_site_index = i
            break
    
    # 如果找到 [DisplaySite] 包含 "Last Contact"
    if display_site_index is not None:
        # 向前搜索所有包含 [Trigger4PinAlignment] 和目標文本的行
        for i in range(display_site_index, len(lines)):
            if "[Trigger4PinAlignment]" in lines[i] and "Wait 2 Sec To Trigger 4 Pin Alignment" in lines[i]:
                # 提取時間戳
                timestamp_part = lines[i].split("[Trigger4PinAlignment]")[0].strip()
                AutoZLog_Last_Trigger4pinalignment_Time = timestamp_part
        
        # 只打印最終結果
        if AutoZLog_Last_Trigger4pinalignment_Time is not None:
            print(f"Found Trigger4PinAlignment timestamp: {AutoZLog_Last_Trigger4pinalignment_Time}")
    
    # 檢查是否找到時間戳
    if AutoZLog_Last_Trigger4pinalignment_Time is None:
        print("Could not find required trigger alignment timestamp in log file.")
    
    return AutoZLog_Last_Trigger4pinalignment_Time

def process_all_txt(file_path, autoz_log_timestamp):
    """處理 ALL.TXT 檔案以提取所需數據，只處理標準值時間戳之後的數據"""
    print(f"Processing ALL.TXT file: {file_path}")
    
    # 檢查 AutoZLog 時間戳是否可用
    if not autoz_log_timestamp:
        print("AutoZLog timestamp not available. Please process AutoZLog.txt first.")
        raise ValueError("AutoZLog timestamp not available")
    
    # 將時間戳從 '2025/04/08 14:59:41.200' 轉換為 '04.08 14:59:41.200'
    converted_timestamp = convert_timestamp_format(autoz_log_timestamp)
    if not converted_timestamp:
        print("Could not convert timestamp format.")
        raise ValueError("Timestamp format conversion failed")
    
    print(f"Searching for events after {converted_timestamp}")
    
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
    
    # 標準值初始化
    x_standard = None
    y_standard = None
    z_standard = None
    standard_timestamp = None
    
    # 在轉換後的時間戳之後找到第一個 ndlp-correct 條目
    found_first_entry = False
    ndlp_pattern = re.compile(r'(\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\.\d{3}).*ndlp-correct: \(\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+) \)')
    
    # 先找到標準值和時間戳
    for line in lines:
        match = ndlp_pattern.search(line)
        if match:
            timestamp = match.group(1)
            
            # 在轉換後的時間戳之後找到第一個 ndlp-correct: 條目
            if timestamp >= converted_timestamp and not found_first_entry:
                standard_timestamp = timestamp
                x_standard = float(match.group(2))
                y_standard = float(match.group(3))
                z_standard = float(match.group(4))
                found_first_entry = True
                print(f"Found standard values at {standard_timestamp}: X={x_standard}, Y={y_standard}, Z={z_standard}")
                break  # 找到標準值後跳出循環
    
    # 確保找到了標準值
    if not found_first_entry:
        print("Could not find standard values after the trigger timestamp.")
        raise ValueError("Standard values not found")
    
    # 第二次遍歷，只處理標準時間戳之後的條目
    for line in lines:
        match = ndlp_pattern.search(line)
        if match:
            timestamp = match.group(1)
            
            # 只處理時間戳大於等於標準時間戳的條目
            if timestamp >= standard_timestamp:
                x_val = float(match.group(2))
                y_val = float(match.group(3))
                z_val = float(match.group(4))
                
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
    
    # 打印提取數據的摘要
    print("Data extraction summary:")
    for wafer_id, data in wafer_ids.items():
        print(f"Wafer ID: {wafer_id}")
        print(f"  Start time: {data['start_time']}")
        print(f"  Number of ndlp-correct entries: {len(data['x_values'])}")
        if data['x_values']:
            print(f"  X value range: {min(data['x_values'])} to {max(data['x_values'])}")
            print(f"  Y value range: {min(data['y_values'])} to {max(data['y_values'])}")
            print(f"  Z value range: {min(data['z_values'])} to {max(data['z_values'])}")
        print("---")
    
    # 返回處理結果
    return {
        'wafer_data': wafer_ids,
        'x_standard': x_standard,
        'y_standard': y_standard,
        'z_standard': z_standard,
        'standard_timestamp': standard_timestamp
    }