import os
import datetime
import wmi
import win32evtlog
from supabase import create_client
import time
import sys
import ctypes
from ctypes import Structure, windll, c_uint, sizeof, byref
import requests
import hashlib
import subprocess

AR_VER = '[v2.2.7]'
SS_DELAY = 3
CPU_USAGE_THRESHOLD = 50
IDLE_TIME_THRESHOLD = 1800
SHUTDOWN_COUNTDOWN = 300
WAIT_HOUR = 18
WAIT_MIN = 30
TARGET_EVENT_IDS = [42, 26, 4001, 109, 1002]

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

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def log_message(message):
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "anex-attendance-record")
    if not os.path.exists(documents_folder):
        os.makedirs(documents_folder)
    log_file_path = os.path.join(documents_folder, "AneX-AR_Log.txt")
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{current_time}] {message}\n")

def check_internet_connection_via_supabase():
    try:
        supabase.table("userlist").select("*").limit(1).execute()
        return True
    except Exception as e:
        log_message(f"嘗試連線至supabase失敗: {e}")
        return False

def wait_for_internet():
    while True:
        if check_internet_connection_via_supabase():
            break
        else:
            log_message("無法連接 Supabase，等待五分鐘後重試...")
            time.sleep(300)

def normalize_mac(mac):
    mac = mac.upper().replace("-", "").replace(":", "").strip("[]")
    formatted_mac = ":".join(mac[i:i+2] for i in range(0, len(mac), 2))
    return f"[{formatted_mac}]"

def get_local_mac_address():
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
    earliest_today_event = None
    for event in events:
        if start_time <= event.TimeGenerated < end_time:
            earliest_today_event = event
            break
    if earliest_today_event:
        log_message(f"本日開機時間為: {earliest_today_event.TimeGenerated}")
        latest_event_time = None
        for event in reversed(events):
            if event.TimeGenerated < earliest_today_event.TimeGenerated:
                latest_event_time = event.TimeGenerated
                break
        if latest_event_time:
            log_message(f"上次關機時間為: {latest_event_time}")
        if latest_event_time and latest_event_time.hour >= 23:
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
    # 將建立表格、啟用 RLS 及新增允許所有存取的 policy 合併為單次執行
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
    
    ALTER TABLE public.{table_name} ENABLE ROW LEVEL SECURITY;
    
    -- 如果已經存在同名 policy，先刪除再重新建立，避免錯誤
    DROP POLICY IF EXISTS "Allow all operations" ON public.{table_name};
    
    CREATE POLICY "Allow all operations"
        ON public.{table_name}
        FOR ALL
        TO public
        USING (true)
        WITH CHECK (true);
    """
    try:
        supabase.rpc("execute_sql", {"sql": create_sql}).execute()
    except Exception as e:
        log_message(f"嘗試建立表格 {table_name} 或套用 RLS/Policy 失敗：{e}")

def fix_sequence(table_name):
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

def get_cpu_usage():
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
    try:
        minutes = delay // 60
        log_message(f"系統將在 {minutes} 分鐘後關機，請盡速儲存您的工作。")
        os.system(f"shutdown /s /f /t {delay}")
    except Exception as e:
        log_message(f"執行 shutdown 時發生錯誤: {e}")

def get_idle_duration():
    """取得使用者鍵盤/滑鼠的閒置秒數"""
    last_input_info = LASTINPUTINFO()
    last_input_info.cbSize = sizeof(last_input_info)
    try:
        if windll.user32.GetLastInputInfo(byref(last_input_info)):
            millis_now = windll.kernel32.GetTickCount()
            last_input_millis = last_input_info.dwTime
            idle_seconds = (millis_now - last_input_millis) / 1000.0
            return idle_seconds
    except Exception as e:
        log_message(f"取得閒置時間時發生錯誤: {e}")
    return 0

def monitor_idle_and_shutdown():
    while True:
        try:
            now = datetime.datetime.now()
            idle_seconds = get_idle_duration()
            cpu_usage_value = get_cpu_usage()
            is_after_office = ((now.hour > WAIT_HOUR) or 
                               (now.hour == WAIT_HOUR and now.minute >= WAIT_MIN))
            cpu_is_idle = (cpu_usage_value <= CPU_USAGE_THRESHOLD)
            mouse_keyboard_idle = (idle_seconds >= IDLE_TIME_THRESHOLD)

            if is_after_office and cpu_is_idle and mouse_keyboard_idle:
                log_message("偵測到時間已晚於18:30、CPU閒置、鍵鼠無操作超過30分鐘，將執行5分鐘倒數關機。")
                shutdown_windows(delay=SHUTDOWN_COUNTDOWN)

            time.sleep(60)
        except Exception as e:
            log_message(f"監控閒置和關機時發生錯誤: {e}")
            time.sleep(60)
            
def check_and_update_anex_ar():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    def get_remote_file_hash(url):
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
        if not os.path.exists(file_path):
            return None
        with open(file_path, "rb") as f:
            content = f.read()
        return hashlib.sha256(content).hexdigest()
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
            break

def check_for_immunity():
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "anex-attendance-record")
    for candidate in ["免死金牌", "免死金牌.txt"]:
        file_path = os.path.join(documents_folder, candidate)
        if os.path.exists(file_path):
            log_message("...發現免死金牌！")
            return True
    return False

def update_attendance_file(employee_data):
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
            while True:
                success = update_attendance_record(datetime.datetime.now().strftime("%Y-%m-%d"), both_data)
                if success:
                    break
                else:
                    log_message("更新出勤記錄時發生錯誤，五分鐘後重試...")
                    time.sleep(300)
            return earliest_today_time, last_shutdown_time

def restart_program(reason):
    log_message(reason)
    subprocess.Popen(["pythonw", LOCAL_FILE])
    sys.exit(0)

if __name__ == "__main__":
    wait_for_internet()
    check_and_update_anex_ar()
    time.sleep(SS_DELAY)
    while True:
        employee_data = fetch_userlist_from_supabase()
        if employee_data:
            break
        log_message("無法從 Supabase 獲取使用者列表，五分鐘後重試...")
        time.sleep(300)
    try:
        earliest_today_time, last_shutdown_time = update_attendance_file(employee_data)
    except Exception as e:
        log_message(f"程序執行失敗: {e}")
        while True:
            time.sleep(300)
            try:
                earliest_today_time, last_shutdown_time = update_attendance_file(employee_data)
                break
            except Exception as e2:
                log_message(f"程序再次執行失敗: {e2}, 五分鐘後重試...")
    if earliest_today_time is not None:
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
                updated_successfully = verify_attendance_record(today_str, employee_info['employee_id'], check_times)
                if updated_successfully:
                    log_message("出勤記錄驗證成功。")
                else:
                    restart_program("出勤記錄驗證失敗，將重新啟動程式...")
            else:
                restart_program("員工資料驗證失敗，將重新啟動程式...")
        else:
            restart_program("本機MAC地址驗證失敗，將重新啟動程式...")
    if check_for_immunity():
        sys.exit(0)
    monitor_idle_and_shutdown()
