import re
from datetime import datetime

def convert_timestamp_format(timestamp):
    """將 AG93000 時間戳記轉換為統一格式 '04.08 14:59:41.200'
    
    輸入格式: '2025/06/06 15:31:05.104'
    輸出格式: '06.06 15:31:05.104'
    """
    try:
        dt = datetime.strptime(timestamp, "%Y/%m/%d %H:%M:%S.%f")
        return f"{dt.month:02d}.{dt.day:02d} {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}.{int(dt.microsecond/1000):03d}"
    except Exception as e:
        print(f"Error converting timestamp: {e}")
        return None

def process_autoz_log(file_path):
    """處理 AG93000 的 AutoZLog.txt 檔案以提取所需數據
    
    新流程：
    1. 從最後一列向上搜尋找到第一個 "[Get Last Contact]"
    2. 從該位置向上搜尋第一個 "[EXEC_INP_CALL]"
    3. 提取 EXEC_INP_CALL 的日期
    4. 提取 Get Last Contact 的時間
    5. 合併為完整時間戳
    
    Returns:
        str: 格式為 '2025/06/06 15:31:05.104' 的完整時間戳，或 None
    """
    # 讀取檔案
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            lines = file.readlines()
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    
    # 正規表達式模式
    # 匹配：[AutoZ Log] 15:31:05.104 Get Last Contact , Last Contact = 26
    get_last_contact_pattern = re.compile(r'\[AutoZ Log\] (\d{2}:\d{2}:\d{2}\.\d{3}) Get Last Contact')
    
    # 匹配：2025-06-06 15:31:01.171 : [EXEC_INP_CALL]
    exec_inp_pattern = re.compile(r'(\d{4}-\d{2}-\d{2}) \d{2}:\d{2}:\d{2}\.\d{3} : \[EXEC_INP_CALL\]')
    
    # 步驟 1：從最後一列向上搜尋找到第一個 "[Get Last Contact]"
    get_last_contact_index = None
    get_last_contact_time = None
    
    for i in range(len(lines)-1, -1, -1):
        match = get_last_contact_pattern.search(lines[i])
        if match:
            get_last_contact_index = i
            get_last_contact_time = match.group(1)  # 提取時間：15:31:05.104
            print(f"Found [Get Last Contact] at line {i+1}")
            break
    
    # 檢查是否找到 Get Last Contact
    if get_last_contact_index is None or get_last_contact_time is None:
        print("Could not find '[Get Last Contact]' in log file.")
        return None
    
    # 步驟 2：從 Get Last Contact 位置向上搜尋第一個 "[EXEC_INP_CALL]"
    exec_inp_date = None
    
    for i in range(get_last_contact_index, -1, -1):
        match = exec_inp_pattern.search(lines[i])
        if match:
            exec_inp_date = match.group(1)  # 提取日期：2025-06-06
            print(f"Found [EXEC_INP_CALL] at line {i+1}")
            break
    
    # 檢查是否找到 EXEC_INP_CALL
    if exec_inp_date is None:
        print("Could not find '[EXEC_INP_CALL]' before '[Get Last Contact]' in log file.")
        return None
    
    # 步驟 3：合併日期和時間
    # 將日期格式從 YYYY-MM-DD 轉換為 YYYY/MM/DD
    exec_inp_date_formatted = exec_inp_date.replace("-", "/")
    
    # 合併為完整時間戳：2025/06/06 15:31:05.104
    full_timestamp = f"{exec_inp_date_formatted} {get_last_contact_time}"
    
    print(f"Successfully constructed timestamp: {full_timestamp}")
    print(f"  Date from EXEC_INP_CALL: {exec_inp_date}")
    print(f"  Time from Get Last Contact: {get_last_contact_time}")
    
    return full_timestamp

def process_all_txt(file_path, autoz_log_timestamp):
    """處理 ALL.TXT 檔案以提取所需數據，只處理 Auto Z complete 點之後的數據"""
    print(f"Processing ALL.TXT file: {file_path}")
    
    # 檢查 Last Contact 時間戳是否可用
    if not autoz_log_timestamp:
        print("Last Contact timestamp not available. Please process AutoZLog.txt first.")
        raise ValueError("Last Contact timestamp not available")
    
    # 將時間戳從 '2025/06/06 15:31:05.104' 轉換為 '06.06 15:31:05.104'
    last_contact_converted = convert_timestamp_format(autoz_log_timestamp)
    if not last_contact_converted:
        print("Could not convert timestamp format.")
        raise ValueError("Timestamp format conversion failed")
    
    print(f"Searching for Auto Z complete point after {last_contact_converted}")
    
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
    
    # 尋找 Auto Z complete 點（Get Last Contact 之後的第一個 ndlp-correct）
    found_autoz_complete = False
    ndlp_pattern = re.compile(r'(\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\.\d{3}).*ndlp-correct: \(\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+),\s*([-+]?\d+\.\d+) \)')
    
    # 第一次遍歷：找到 Auto Z complete 點
    for line in lines:
        match = ndlp_pattern.search(line)
        if match:
            timestamp = match.group(1)
            
            # Get Last Contact 之後的第一個 ndlp-correct 就是 Auto Z complete 點
            if timestamp >= last_contact_converted and not found_autoz_complete:
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
        print("Could not find Auto Z complete point (first ndlp-correct after Get Last Contact).")
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