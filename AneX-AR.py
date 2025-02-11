import os
import datetime
import wmi
import win32evtlog
from supabase import create_client
import time
import sys
import ctypes
from ctypes import Structure, windll, c_uint, sizeof, byref
from ctypes import wintypes
import requests
import hashlib
import subprocess

# === 版本與常數設定 ===
AR_VER = '[v2.2.5]'
SS_DELAY = 5
CPU_USAGE_THRESHOLD = 20
IDLE_TIME_THRESHOLD = 18
SHUTDOWN_COUNTDOWN = 300
WAIT_HOUR = 18
WAIT_MIN = 30
TARGET_EVENT_IDS = [42, 26, 4001, 109, 1002]

# === 檔案與路徑設定 ===
base_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(base_dir, "config.env")

while not os.path.exists(env_path):
    print(f"配置文件不存在：{env_path}，將在五分鐘後重試...")
    time.sleep(300)

def load_env(filepath):
    """
    從指定路徑讀取環境變數檔案 (config.env)，並回傳字典
    """
    env_vars = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            key, value = line.split("=", 1)
            env_vars[key.strip()] = value.strip()
    return env_vars

while True:
    env = load_env(env_path)
    GITHUB_RAW_URL = env.get("GITHUB_URL")
    SUPABASE_URL = env.get("SUPABASE_URL")
    SUPABASE_KEY = env.get("SUPABASE_KEY")
    LOCAL_FILE = env.get("LOCAL_FILE")
    if not SUPABASE_URL or not SUPABASE_KEY:
        print("Supabase 相關設定不完整，五分鐘後重試...")
        time.sleep(300)
        continue
    break

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def log_message(message):
    """
    將訊息記錄到「我的文件」下的 anex-attendance-record 資料夾中，
    並在檔案 AneX-AR_Log.txt 追加寫入
    """
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "anex-attendance-record")
    if not os.path.exists(documents_folder):
        os.makedirs(documents_folder)
    log_file_path = os.path.join(documents_folder, "AneX-AR_Log.txt")
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{current_time}] {message}\n")

def check_internet_connection_via_supabase():
    """
    嘗試透過對 Supabase 執行最簡單的查詢來驗證網路連線
    如果失敗，會拋出例外
    """
    try:
        supabase.table("userlist").select("*").limit(1).execute()
        return True
    except Exception as e:
        log_message(f"嘗試連線至supabase失敗: {e}")
        return False

def wait_for_internet():
    """
    不斷檢查是否能連上 Supabase，如果無法連線就等待五分鐘再重試
    """
    while True:
        if check_internet_connection_via_supabase():
            break
        else:
            log_message("無法連接 Supabase，等待五分鐘後重試...")
            time.sleep(300)

def normalize_mac(mac):
    """
    將 MAC 轉成標準格式，例如 AA:BB:CC:DD:EE:FF
    """
    mac = mac.upper().replace("-", "").replace(":", "").strip("[]")
    formatted_mac = ":".join(mac[i:i+2] for i in range(0, len(mac), 2))
    return f"[{formatted_mac}]"

def get_local_mac_address():
    """
    取得本機第一張啟用的網路卡的 MAC 位址並標準化
    """
    try:
        c = wmi.WMI()
        net_adapters = c.Win32_NetworkAdapterConfiguration(IPEnabled=1)
        if net_adapters:
            localMAC = net_adapters[0].MACAddress
            return normalize_mac(localMAC) if localMAC else None
        else:
            log_message("找不到任何已啟用的網路卡。")
        return None
    except Exception as e:
        log_message(f"取得本機 MAC 位址時發生錯誤: {e}")
        return None

def get_today_and_previous_events():
    """
    讀取系統日誌 (System Log)，取得今日最早開機事件時間 (earliest_today_event)
    以及前一次關機時間 (latest_event_time)
    如果今日開機事件為 None，代表失敗
    """
    server = 'localhost'
    log_type = 'System'
    try:
        handle = win32evtlog.OpenEventLog(server, log_type)
    except Exception as e:
        log_message(f"無法打開 System Log 事件紀錄: {e}")
        return None, None

    today = datetime.datetime.now()
    start_time = datetime.datetime(today.year, today.month, today.day)
    end_time = start_time + datetime.timedelta(days=1)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

    events = []
    while True:
        records = win32evtlog.ReadEventLog(handle, flags, 0)
        if not records:
            break
        events.extend(records)

    win32evtlog.CloseEventLog(handle)

    if not events:
        log_message("未找到任何系統事件")
        return None, None

    events.sort(key=lambda e: e.TimeGenerated)

    # 找今日最早事件
    earliest_today_event = None
    for event in events:
        if start_time <= event.TimeGenerated < end_time:
            earliest_today_event = event
            break

    if earliest_today_event:
        log_message(f"本日開機時間為: {earliest_today_event.TimeGenerated}")
        latest_event_time = None
        # 向前找(逆序)最晚的事件時間
        for event in reversed(events):
            if event.TimeGenerated < earliest_today_event.TimeGenerated:
                latest_event_time = event.TimeGenerated
                break

        if latest_event_time:
            log_message(f"上次關機時間為: {latest_event_time}")

        # 如果發現上一次關機時間在22點之後，表示昨晚未準時關機
        # 嘗試再查找當天 18:30 ~ 23:59:59 之間是否有符合條件的事件
        if latest_event_time and latest_event_time.hour >= 22:
            log_message("昨天未準時關機，開始檢查是否有符合條件的事件。")
            target_date = latest_event_time.date()
            event_start = datetime.datetime.combine(target_date, datetime.time(18, 30))
            event_end = datetime.datetime.combine(target_date, datetime.time(23, 59, 59))

            earliest_event_time = None
            for ev in events:
                if event_start <= ev.TimeGenerated <= event_end:
                    if (ev.EventID & 0xFFFF) in TARGET_EVENT_IDS:
                        if earliest_event_time is None or ev.TimeGenerated < earliest_event_time:
                            earliest_event_time = ev.TimeGenerated
                            log_message(f"找到事件: {ev.EventID}，時間: {earliest_event_time}")

            if earliest_event_time:
                latest_event_time = earliest_event_time
                log_message(f"將昨天下班時間更新為符合條件的事件時間: {latest_event_time}")
            else:
                log_message("未找到符合條件的事件，保持原始的昨天下班時間。")

        return earliest_today_event.TimeGenerated, latest_event_time
    else:
        log_message("未找到今日的最早事件，無法確定上班時間。")
        return None, None

def fetch_userlist_from_supabase():
    """
    從 Supabase 取得 userlist 表的所有資料
    如果取得失敗，回傳 None
    """
    try:
        response = supabase.table('userlist').select("*").execute()
        if response.data:
            return response.data
        else:
            log_message("從 Supabase 獲取使用者列表失敗，數據為空。請檢查 Supabase 配置或數據表。")
            return None
    except Exception as e:
        log_message(f"無法從 Supabase 獲取使用者列表：{e}")
        return None

def ensure_daily_table_exists(table_name):
    """
    如果指定的考勤表不存在，就建立
    """
    create_sql = f"""
    CREATE TABLE IF NOT EXISTS public.{table_name} (
        id BIGSERIAL PRIMARY KEY,
        date VARCHAR(10),
        employee_id VARCHAR(50),
        employee_name VARCHAR(100),
        mac_address VARCHAR(50),
        check_in TIMESTAMP NULL,
        check_out TIMESTAMP NULL,
        version VARCHAR(20)
    );
    """
    try:
        supabase.rpc("execute_sql", {"sql": create_sql}).execute()
    except Exception as e:
        log_message(f"嘗試建立表格 {table_name} 失敗：{e}")

def fix_sequence(table_name):
    """
    修正表格的序列 (id BIGSERIAL) 使其位於正確的值
    """
    try:
        max_id_sql = f"SELECT MAX(id) AS max_id FROM public.{table_name}"
        max_id_resp = supabase.rpc("execute_sql", {"sql": max_id_sql}).execute()
        if max_id_resp.data:
            current_max_id = max_id_resp.data[0].get("max_id") or 0
            setval_sql = f"SELECT setval(pg_get_serial_sequence('public.{table_name}', 'id'), {current_max_id + 1});"
            supabase.rpc("execute_sql", {"sql": setval_sql}).execute()
            log_message(f"已將序列調整為 {current_max_id + 1}")
    except Exception as e:
        log_message(f"修正序列時發生錯誤: {e}")

def update_attendance_record(date_str, employee_data):
    """
    更新指定員工當天的考勤紀錄 (包含上班時間 check_in、下班時間 check_out)
    成功回傳 True，失敗回傳 False
    """
    table_name = "attendance_" + date_str.replace('-', '_')
    try:
        ensure_daily_table_exists(table_name)
        check_in_val = employee_data.get('check_in') or None
        check_out_val = employee_data.get('check_out') or None
        attendance_data = {
            'date': date_str,
            'employee_id': employee_data['employee_id'],
            'employee_name': employee_data['employee_name'],
            'mac_address': employee_data['mac_address'],
            'check_in': check_in_val,
            'check_out': check_out_val,
            'version': AR_VER
        }
        existing = supabase.table(table_name).select("*").eq('date', date_str).eq('employee_id', employee_data['employee_id']).execute()
        if existing.data:
            supabase.table(table_name).update(attendance_data).eq('date', date_str).eq('employee_id', employee_data['employee_id']).execute()
        else:
            try:
                supabase.table(table_name).insert(attendance_data).execute()
            except Exception as e:
                err_str = str(e)
                if "23505" in err_str:
                    log_message(f"偵測到主鍵衝突，錯誤內容：{e}")
                    same_mac = supabase.table(table_name).select("id").eq('mac_address', attendance_data['mac_address']).execute()
                    if same_mac.data:
                        log_message("資料表中已有相同 MAC，略過重新插入與修正序列。")
                    else:
                        log_message("無相同 MAC，嘗試修正序列。")
                        fix_sequence(table_name)
                else:
                    raise
        return True
    except Exception as e:
        log_message(f"更新出勤記錄失敗：{e}")
        return False

def verify_attendance_record(date_str, employee_id, check_times=None):
    """
    驗證當天指定員工的考勤紀錄，若 check_times 有值，則檢查 check_in/check_out 是否對應
    成功回傳 True，否則 False
    """
    table_name = "attendance_" + date_str.replace('-', '_')
    try:
        response = supabase.table(table_name).select("*").eq('date', date_str).eq('employee_id', employee_id).execute()
        if response.data:
            record = response.data[0]
            if check_times:
                check_in_str = str(record.get('check_in')).replace('T', ' ')
                check_out_str = str(record.get('check_out')).replace('T', ' ')
                times_in_record = [check_in_str, check_out_str]
                for check_time in check_times:
                    if check_time in times_in_record:
                        return True
                return False
            return True
        else:
            log_message("未查詢到任何紀錄。")
        return False
    except Exception as e:
        log_message(f"驗證出勤紀錄失敗：{e}")
        return False

class LASTINPUTINFO(Structure):
    _fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_uint)
    ]

def get_idle_duration():
    """
    取得鍵鼠閒置的秒數
    """
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = sizeof(lastInputInfo)
    try:
        if windll.user32.GetLastInputInfo(byref(lastInputInfo)):
            millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
            return millis / 1000.0
    except Exception as e:
        log_message(f"取得閒置時間時發生錯誤: {e}")
    return 0

def get_cpu_usage():
    """
    取得 CPU 即時使用率
    """
    try:
        c = wmi.WMI()
        cpus = c.Win32_Processor()
        if cpus:
            usage_list = [cpu.LoadPercentage for cpu in cpus if cpu.LoadPercentage is not None]
            if usage_list:
                return round(sum(usage_list) / len(usage_list), 2)
        return 0
    except Exception as e:
        log_message(f"取得CPU使用率時發生錯誤: {e}")
        return 0

def shutdown_windows(delay=300):
    """
    執行系統關機並提示倒數
    """
    try:
        log_message(f"系統將在 {delay // 60} 分鐘後關機，請儲存工作。")
        os.system(f"shutdown /s /f /t {delay}")
    except Exception as e:
        log_message(f"執行 shutdown 時發生錯誤: {e}")

def monitor_idle_and_shutdown():
    """
    監控下班後的閒置時間，若達到條件，進行 5 分鐘倒數關機
    """
    has_logged_start_message = False
    while True:
        try:
            now = datetime.datetime.now()
            # 若時間已到下班後
            if (now.hour > WAIT_HOUR) or (now.hour == WAIT_HOUR and now.minute >= WAIT_MIN):
                if not has_logged_start_message:
                    log_message("已到下班時間，開始監控閒置狀態")
                    has_logged_start_message = True

                idle_seconds = get_idle_duration()
                cpu_usage_value = get_cpu_usage() or 0

                if idle_seconds >= IDLE_TIME_THRESHOLD and cpu_usage_value <= CPU_USAGE_THRESHOLD:
                    log_message("偵測到鍵鼠閒置超過設定時間，將執行 5 分鐘倒數關機。")
                    shutdown_windows(delay=SHUTDOWN_COUNTDOWN)
            else:
                has_logged_start_message = False

            time.sleep(60)

        except Exception as e:
            log_message(f"監控閒置和關機時發生錯誤: {e}")
            time.sleep(60)

def check_and_update_anex_ar():
    """
    檢查遠端檔案雜湊 (GITHUB_RAW_URL)，若有新版本就下載替換本地檔案並重新啟動
    如果失敗，則每五分鐘重試一次
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    def get_remote_file_hash(url):
        """
        取得遠端檔案的 SHA256，並回傳 (hash, content)
        如果取得失敗，回傳 (None, None)
        """
        try:
            response = requests.get(url)
            if response.status_code == 200:
                content = response.content
                return hashlib.sha256(content).hexdigest(), content
            else:
                log_message(f"無法取得遠端檔案，HTTP 狀態碼: {response.status_code}")
                return None, None
        except Exception as e:
            log_message(f"取得遠端檔案時發生錯誤: {e}")
            return None, None

    def get_local_file_hash(file_path):
        """
        取得本地檔案的 SHA256
        """
        if not os.path.exists(file_path):
            return None
        with open(file_path, "rb") as f:
            content = f.read()
        return hashlib.sha256(content).hexdigest()

    # === 不斷重試，直到成功取得遠端檔案 (或者雜湊比對成功) 為止 ===
    while True:
        remote_hash, remote_content = get_remote_file_hash(GITHUB_RAW_URL)
        if remote_hash is None:
            log_message("遠端檔案無法取得，五分鐘後重試...")
            time.sleep(300)
            continue

        local_hash = get_local_file_hash(LOCAL_FILE)
        if local_hash != remote_hash:
            log_message("偵測到更新，正在下載...")
            try:
                with open(LOCAL_FILE, "wb") as f:
                    f.write(remote_content)
                log_message("更新完成！")
                # 以 pythonw 重新執行新版程式，然後結束當前程式
                subprocess.Popen(["pythonw", LOCAL_FILE])
                sys.exit(0)
            except PermissionError as e:
                log_message(f"無法寫入檔案 {LOCAL_FILE}: {e}")
                log_message("五分鐘後重試...")
                time.sleep(300)
            except Exception as e:
                log_message(f"下載檔案時發生未知錯誤: {e}")
                log_message("五分鐘後重試...")
                time.sleep(300)
        else:
            log_message(f"程式啟動 {AR_VER}")
            break  # 雜湊相同，無需更新，離開迴圈

def check_for_immunity():
    """
    檢查使用者的 Documents/anex-attendance-record 是否存在「免死金牌」或「免死金牌.txt」
    有的話就回傳 True
    """
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "anex-attendance-record")
    for candidate in ["免死金牌", "免死金牌.txt"]:
        file_path = os.path.join(documents_folder, candidate)
        if os.path.exists(file_path):
            log_message("...發現免死金牌！")
            return True
    return False

def update_attendance_file(employee_data):
    """
    透過讀取系統日誌取得今日開機與上次關機時間，並更新出勤資料表
    如果找不到今日開機 (earliest_today_time = None)，不斷重試直到成功
    如果取得 MAC 或員工資料失敗，也不斷重試
    """
    while True:
        earliest_today_time, last_shutdown_time = get_today_and_previous_events()
        if earliest_today_time is None:
            log_message("無法取得今日最早開機時間，五分鐘後重試...")
            time.sleep(300)
            continue

        while True:
            mac_address = get_local_mac_address()
            if not mac_address:
                log_message("無法取得本機MAC地址，五分鐘後重試...")
                time.sleep(300)
                continue
            employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
            if not employee_info:
                log_message(f"找不到匹配的員工資料 (MAC={mac_address})，五分鐘後重試...")
                time.sleep(300)
                continue

            log_message(f"{employee_info['employee_id']} {employee_info['name']} {employee_info['mac_address']}")
            both_data = {
                'employee_id': employee_info['employee_id'],
                'employee_name': employee_info['name'],
                'mac_address': mac_address,
                'check_in': str(earliest_today_time),
                'check_out': str(last_shutdown_time) if last_shutdown_time else None
            }

            # 不斷嘗試更新出勤紀錄，成功才離開
            while True:
                success = update_attendance_record(datetime.datetime.now().strftime("%Y-%m-%d"), both_data)
                if success:
                    break
                else:
                    log_message("更新出勤記錄時發生錯誤，五分鐘後重試...")
                    time.sleep(300)

            # 成功更新，就離開這個 while
            return earliest_today_time, last_shutdown_time

if __name__ == "__main__":
    # 1. 等待網路連線穩定
    wait_for_internet()

    # 2. 檢查並更新程式
    check_and_update_anex_ar()

    # 3. 延遲 SS_DELAY 秒再繼續
    time.sleep(SS_DELAY)

    # 4. 取得員工資料 (userlist)，若失敗就五分鐘後重試
    while True:
        employee_data = fetch_userlist_from_supabase()
        if employee_data:
            break
        log_message("無法從 Supabase 獲取使用者列表，五分鐘後重試...")
        time.sleep(300)

    # 5. 更新本機考勤記錄 (直到成功才會離開函式)
    try:
        earliest_today_time, last_shutdown_time = update_attendance_file(employee_data)
    except Exception as e:
        log_message(f"程序執行失敗: {e}")
        # 失敗也五分鐘後重試，不讓整個程式結束
        while True:
            time.sleep(300)
            try:
                earliest_today_time, last_shutdown_time = update_attendance_file(employee_data)
                break
            except Exception as e2:
                log_message(f"程序再次執行失敗: {e2}, 五分鐘後重試...")

    # 6. 驗證考勤是否正確 (不斷重試至成功)
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    mac_address = get_local_mac_address()
    if mac_address:
        employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
        if employee_info:
            check_times = []
            if earliest_today_time:
                check_times.append(str(earliest_today_time))
            if last_shutdown_time:
                check_times.append(str(last_shutdown_time))

            while True:
                updated_successfully = verify_attendance_record(today_str, employee_info['employee_id'], check_times)
                if updated_successfully:
                    log_message("出勤記錄驗證成功。")
                    break
                else:
                    log_message("出勤記錄驗證失敗，五分鐘後重試...")
                    time.sleep(300)
        else:
            log_message("驗證失敗，無法匹配員工資料。五分鐘後重試...")
            while True:
                time.sleep(300)
                mac_address = get_local_mac_address()
                if mac_address:
                    employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
                    if employee_info:
                        check_times = []
                        if earliest_today_time:
                            check_times.append(str(earliest_today_time))
                        if last_shutdown_time:
                            check_times.append(str(last_shutdown_time))
                        updated_successfully = verify_attendance_record(today_str, employee_info['employee_id'], check_times)
                        if updated_successfully:
                            log_message("出勤記錄驗證成功。")
                            break
                        else:
                            log_message("出勤記錄驗證仍然失敗，繼續五分鐘後重試...")
                    else:
                        log_message("還是無法匹配員工資料，再次五分鐘後重試...")
                else:
                    log_message("依然無法取得 MAC 地址，再次五分鐘後重試...")
    else:
        # 若連 MAC 都取不到，也要五分鐘後重試
        log_message("無法取得本機MAC地址，五分鐘後重試取得 MAC。")
        while True:
            time.sleep(300)
            mac_address = get_local_mac_address()
            if mac_address:
                employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
                if employee_info:
                    check_times = []
                    if earliest_today_time:
                        check_times.append(str(earliest_today_time))
                    if last_shutdown_time:
                        check_times.append(str(last_shutdown_time))
                    updated_successfully = verify_attendance_record(today_str, employee_info['employee_id'], check_times)
                    if updated_successfully:
                        log_message("出勤記錄驗證成功。")
                        break
                    else:
                        log_message("出勤記錄驗證失敗，再次五分鐘後重試...")
                else:
                    log_message("還是無法匹配員工資料，再次五分鐘後重試...")
            else:
                log_message("依然無法取得 MAC 地址，再次五分鐘後重試...")

    # 7. 檢查免死金牌
    if check_for_immunity():
        # 若找到免死金牌，就結束程式 (不進入監控)
        sys.exit(0)

    # 8. 進入閒置監控 (不會結束程式)
    monitor_idle_and_shutdown()
