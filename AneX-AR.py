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

AR_VER = '[v2.2.4]'
SS_DELAY = 5                  # 延遲啟動(秒)
CPU_USAGE_THRESHOLD = 20      # CPU使用率門檻(%)
IDLE_TIME_THRESHOLD = 1800    # 鍵鼠無操作時間(秒) = 30 分鐘
SHUTDOWN_COUNTDOWN = 300      # 關機前倒數時間(秒) = 5 分鐘
WAIT_HOUR = 18                # 檢測閒置開始時間(時)
WAIT_MIN = 30                 # 檢測閒置開始時間(分)
TARGET_EVENT_IDS = [42, 26, 4001, 109, 1002]  # 未關機事件檢查
LOCAL_FILE = "AneX-AR.py"

base_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(base_dir, "config.env")
if not os.path.exists(env_path):
    print(f"配置文件不存在：{env_path}")
    sys.exit(1)
def load_env(filepath):
    env_vars = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            key, value = line.split("=", 1)
            env_vars[key.strip()] = value.strip()
    return env_vars

env = load_env(env_path)
GITHUB_RAW_URL = env.get("GITHUB_URL")
SUPABASE_URL = env.get("SUPABASE_URL")
SUPABASE_KEY = env.get("SUPABASE_KEY")
LOCAL_FILE = env.get("LOCAL_FILE")

if not SUPABASE_URL or not SUPABASE_KEY:
    sys.exit(1)
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# --------------------------------------------------
# 日誌紀錄
# --------------------------------------------------
def log_message(message):
    """
    將日誌紀錄到使用者 Documents/anex-attendance-record/AneX-AR_Log.txt 中。
    """
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "anex-attendance-record")
    if not os.path.exists(documents_folder):
        os.makedirs(documents_folder)
    log_file_path = os.path.join(documents_folder, "AneX-AR_Log.txt")
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{current_time}] {message}\n")

# --------------------------------------------------
# 檢查網路連線
# --------------------------------------------------
def check_internet_connection_via_supabase():
    """
    嘗試連線至 Supabase 以檢查網路是否正常
    """
    try:
        supabase.table("userlist").select("*").limit(1).execute()
        return True
    except Exception as e:
        log_message(f"嘗試連線至supabase失敗: {e}")
        return False

def wait_for_internet():
    """
    不斷嘗試連接 Supabase，若失敗則等待五分鐘後重試
    """
    while True:
        if check_internet_connection_via_supabase():
            break
        else:
            log_message("無法連接 Supabase，等待五分鐘後重試...")
            time.sleep(300)

# --------------------------------------------------
# MAC 相關
# --------------------------------------------------
def normalize_mac(mac):
    mac = mac.upper().replace("-", "").replace(":", "").strip("[]")
    formatted_mac = ":".join(mac[i:i+2] for i in range(0, len(mac), 2))
    return f"[{formatted_mac}]"

def get_local_mac_address():
    """
    取得本機的 MAC 位址，若不存在則回傳 None
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

# --------------------------------------------------
# Windows 事件紀錄
# --------------------------------------------------
def get_today_and_previous_events():
    """
    取得今日最早開機事件與前一次關機事件，並進行部分時間判斷與處理
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
    read_count = 0
    while True:
        records = win32evtlog.ReadEventLog(handle, flags, 0)
        if not records:
            break
        events.extend(records)
        read_count += len(records)

    win32evtlog.CloseEventLog(handle)

    if not events:
        log_message("未找到任何系統事件")
        return None, None

    events.sort(key=lambda e: e.TimeGenerated)

    # 取得今日最早事件
    earliest_today_event = None
    for event in events:
        if start_time <= event.TimeGenerated < end_time:
            earliest_today_event = event
            break

    if earliest_today_event:
        log_message(f"本日開機時間為: {earliest_today_event.TimeGenerated}")
        # 往回找離這個事件最近的一筆事件 (作為上次下班時間)
        latest_event_time = None
        for event in reversed(events):
            if event.TimeGenerated < earliest_today_event.TimeGenerated:
                latest_event_time = event.TimeGenerated
                break

        if latest_event_time:
            log_message(f"上次關機時間為: {latest_event_time}")

        # 如果昨天下班時間超過晚上 22:00，檢查指定的 Event ID
        if latest_event_time and latest_event_time.hour >= 22:
            log_message("昨天未準時關機，開始檢查是否有符合條件的事件。")
            target_date = latest_event_time.date()
            event_start = datetime.datetime.combine(target_date, datetime.time(18, 30))
            event_end = datetime.datetime.combine(target_date, datetime.time(23, 59, 59))

            earliest_event_time = None
            for ev in events:
                if event_start <= ev.TimeGenerated <= event_end:
                    # 注意 (ev.EventID & 0xFFFF) 以取得真實的 Event ID
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

# --------------------------------------------------
# Supabase 相關
# --------------------------------------------------
def fetch_userlist_from_supabase():
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
    try:
        # 先取得表中最大 id
        max_id_sql = f"SELECT MAX(id) AS max_id FROM public.{table_name}"
        max_id_resp = supabase.rpc("execute_sql", {"sql": max_id_sql}).execute()
        if max_id_resp.data:
            current_max_id = max_id_resp.data[0].get("max_id") or 0
            # 將 sequence 調整為 (current_max_id + 1)
            setval_sql = (
                f"SELECT setval("
                f"pg_get_serial_sequence('public.{table_name}', 'id'), "
                f"{current_max_id + 1}"
                f");"
            )
            supabase.rpc("execute_sql", {"sql": setval_sql}).execute()
            log_message(f"已將序列調整為 {current_max_id + 1}")
    except Exception as e:
        log_message(f"修正序列時發生錯誤: {e}")

def update_attendance_record(date_str, employee_data):
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

        existing = supabase.table(table_name).select("*") \
            .eq('date', date_str) \
            .eq('employee_id', employee_data['employee_id']) \
            .execute()

        if existing.data:
            supabase.table(table_name).update(attendance_data) \
                .eq('date', date_str) \
                .eq('employee_id', employee_data['employee_id']) \
                .execute()
        else:
            try:
                supabase.table(table_name).insert(attendance_data).execute()
            except Exception as e:
                err_str = str(e)
                if "23505" in err_str:
                    log_message(f"偵測到主鍵衝突，錯誤內容：{e}")
                    same_mac = supabase.table(table_name).select("id") \
                        .eq('mac_address', attendance_data['mac_address']) \
                        .execute()
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
    table_name = "attendance_" + date_str.replace('-', '_')
    try:
        response = supabase.table(table_name).select("*") \
            .eq('date', date_str) \
            .eq('employee_id', employee_id) \
            .execute()

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

# --------------------------------------------------
# 鍵鼠與 CPU 閒置檢查
# --------------------------------------------------
class LASTINPUTINFO(Structure):
    _fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_uint)
    ]

def get_idle_duration():
    """
    取得使用者鍵鼠閒置時間（秒）
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
    取得 CPU 使用率（%）
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

# --------------------------------------------------
# 開機、關機時間紀錄
# --------------------------------------------------
def update_attendance_file(employee_data):
    """
    更新本日上下班記錄
    """
    earliest_today_time, last_shutdown_time = get_today_and_previous_events()
    if earliest_today_time is None:
        log_message("無法取得今日最早開機時間，無法更新出勤紀錄。")
        return None, None

    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    mac_address = get_local_mac_address()
    if not mac_address:
        log_message("無法取得本機MAC地址，請檢查網路配置。")
        return None, None

    employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
    if not employee_info:
        log_message(f"找不到匹配的員工資料 (MAC={mac_address})，請檢查 userlist 資料。")
        return None, None

    log_message(f"{employee_info['employee_id']} {employee_info['name']} {employee_info['mac_address']}")

    both_data = {
        'employee_id': employee_info['employee_id'],
        'employee_name': employee_info['name'],
        'mac_address': mac_address,
        'check_in': str(earliest_today_time),
        'check_out': str(last_shutdown_time) if last_shutdown_time else None
    }

    success = update_attendance_record(today_str, both_data)
    if not success:
        log_message("更新出勤記錄時發生錯誤。")

    return earliest_today_time, last_shutdown_time

# --------------------------------------------------
# 關機動作
# --------------------------------------------------
def shutdown_windows(delay=300):
    """
    顯示對話框提示用戶後，呼叫系統 shutdown 計時關機
    """
    try:
        message = f"系統將在 {delay // 60} 分鐘後關機，請儲存工作！"
        user_response = ctypes.windll.user32.MessageBoxW(0, message, "系統關機通知", 1)
        if user_response == 1:  # 如果用戶點擊「確定」
            os.system(f"shutdown /s /f /t {delay}")
            log_message(f"已呼叫系統內建 shutdown 命令，倒數 {delay} 秒後關機。")
        else:
            log_message("用戶取消了關機操作。")
    except Exception as e:
        log_message(f"執行 shutdown 時發生錯誤: {e}")

def monitor_idle_and_shutdown():
    """
    進入迴圈，不斷檢查閒置狀態，
    若超過閒置時間且 CPU 使用率低，則執行關機動作
    """
    while True:
        try:
            now = datetime.datetime.now()
            if (now.hour > WAIT_HOUR) or (now.hour == WAIT_HOUR and now.minute >= WAIT_MIN):
                idle_seconds = get_idle_duration()
                cpu_usage_value = get_cpu_usage() or 0
                log_message(f"檢查閒置狀態 => Idle: {idle_seconds:.0f} 秒, CPU: {cpu_usage_value}%")

                if idle_seconds >= IDLE_TIME_THRESHOLD and cpu_usage_value <= CPU_USAGE_THRESHOLD:
                    log_message("偵測到鍵鼠閒置超過設定時間，將執行 5 分鐘倒數關機。")
                    shutdown_windows(delay=SHUTDOWN_COUNTDOWN)
            time.sleep(60)
        except Exception as e:
            log_message(f"監控閒置和關機時發生錯誤: {e}")
            time.sleep(60)

# --------------------------------------------------
# 新增：檢查並更新 AneX-AR.py 的函式
# --------------------------------------------------
def check_and_update_anex_ar():
    """
    檢查 GitHub 上的 AneX-AR.py 是否有更新，若有就下載，
    下載後以無視窗方式執行，並結束當前程式。
    若無更新則直接返回，讓主程式繼續往下走。
    """
    # 先切到本程式所在目錄，避免路徑問題
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # 取得遠端檔案的內容與雜湊
    def get_remote_file_hash(url):
        response = requests.get(url)
        if response.status_code == 200:
            content = response.content
            return hashlib.sha256(content).hexdigest(), content
        else:
            log_message(f"無法取得遠端檔案，HTTP 狀態碼: {response.status_code}")
            return None, None

    # 取得本地檔案的雜湊
    def get_local_file_hash(file_path):
        if not os.path.exists(file_path):
            return None
        with open(file_path, "rb") as f:
            content = f.read()
        return hashlib.sha256(content).hexdigest()

    # 實際檢查更新
    log_message("開始檢查 AneX-AR.py 是否有更新...")

    remote_hash, remote_content = get_remote_file_hash(GITHUB_RAW_URL)
    if remote_hash is None:
        log_message("遠端檔案無法取得，放棄更新。")
        return  # 無法檢查更新，直接返回

    local_hash = get_local_file_hash(LOCAL_FILE)
    if local_hash != remote_hash:
        log_message("偵測到 AneX-AR.py 有更新，正在下載...")
        try:
            with open(LOCAL_FILE, "wb") as f:
                f.write(remote_content)
            log_message("AneX-AR.py 已更新完成，開始以無視窗方式執行新版本...")

            # 以無視窗方式啟動 AneX-AR.py
            # 如果你希望它在背景執行，可以使用 pythonw
            # 若使用 python，則需搭配 STARTUPINFO 隱藏視窗
            # 此處使用 pythonw 最簡單
            subprocess.Popen(["pythonw", LOCAL_FILE])
            log_message("新版本 AneX-AR.py 已啟動，現在結束 main.py。")
            sys.exit(0)  # 結束當前程式
        except PermissionError as e:
            log_message(f"無法寫入檔案 {LOCAL_FILE}，請檢查檔案是否被鎖定或是否有寫入權限。錯誤: {e}")
        except Exception as e:
            log_message(f"下載檔案時發生未知錯誤: {e}")
    else:
        log_message("AneX-AR.py 已是最新版本，無需更新。")

# --------------------------------------------------
# 主程式入口
# --------------------------------------------------
if __name__ == "__main__":
    # 1. 先檢查網路
    wait_for_internet()

    # 2. 檢查 AneX-AR.py 是否需要更新，若需要則更新並以無視窗方式執行後結束
    check_and_update_anex_ar()

    # 如果走到這裡，代表沒有更新，繼續執行後續流程
    log_message(f"程式啟動 {AR_VER}")
    time.sleep(SS_DELAY)

    try:
        employee_data = fetch_userlist_from_supabase()
        if not employee_data:
            log_message("無法從 Supabase 獲取使用者列表，請檢查網路連線或 Supabase 資料庫。程序即將退出。")
            sys.exit(0)
        earliest_today_time, last_shutdown_time = update_attendance_file(employee_data)
    except Exception as e:
        log_message(f"程序執行失敗: {e}")
        sys.exit(0)

    if earliest_today_time is not None:
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        mac_address = get_local_mac_address()
        if not mac_address:
            log_message("無法取得本機MAC地址，程序即將退出。請檢查網路配置。")
            sys.exit(0)

        employee_info = next((emp for emp in employee_data if emp['mac_address'] == mac_address), None)
        if employee_info:
            check_times = [str(earliest_today_time), str(last_shutdown_time) if last_shutdown_time else None]
            check_times = [time for time in check_times if time]

            updated_successfully = verify_attendance_record(today_str, employee_info['employee_id'], check_times)
            if updated_successfully:
                log_message("出勤記錄驗證成功。")
            else:
                log_message("出勤記錄驗證失敗。")
        else:
            log_message("驗證失敗，無法匹配員工資料。")

    # 3. 監控閒置，並於特定時間後關機
    monitor_idle_and_shutdown()
