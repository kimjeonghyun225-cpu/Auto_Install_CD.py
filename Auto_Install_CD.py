import os
import time
import subprocess
import json
import sys
import warnings

warnings.filterwarnings(
    "ignore",
    message="Unable to find acceptable character detection dependency.*"
)

import requests
import base64
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from os.path import basename, join, isdir, abspath, getmtime, splitext
from threading import Lock

try:
    import msvcrt
except ImportError:
    msvcrt = None

COLOR_RESET = "\033[0m"
COLOR_CYAN = "\033[96m"
COLOR_GREEN = "\033[92m"
COLOR_YELLOW = "\033[93m"
COLOR_RED = "\033[91m"

# ==========================================
# ⚙️ 1. 설정 및 경로 관리
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config_qa_installer.json")
TARGET_EXTENSIONS = ('.apk', '.bat', '.obb')
TEMP_DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "Downloads")

if not os.path.exists(TEMP_DOWNLOAD_DIR):
    os.makedirs(TEMP_DOWNLOAD_DIR)

def normalize_input_path(path_text):
    cleaned = path_text.strip().replace('"', '').replace("'", "")
    if not cleaned:
        return ""
    return os.path.abspath(os.path.expandvars(cleaned))

def sanitize_display_path(path_text):
    if not path_text:
        return ""
    cleaned = str(path_text).replace("\\\\?\\", "")
    return os.path.normpath(cleaned)

def is_valid_base_path(path_text):
    return bool(path_text) and os.path.exists(path_text) and isdir(path_text)

def load_config_data():
    if not os.path.exists(CONFIG_FILE):
        return {}

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}

def save_config_data(data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def update_path_history(data, path_text):
    history = data.get("path_history", [])
    if not isinstance(history, list):
        history = []

    normalized_path = normalize_input_path(path_text)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history = [entry for entry in history if isinstance(entry, dict) and entry.get("path") != normalized_path]
    history.insert(0, {"path": normalized_path, "saved_at": timestamp})
    data["path_history"] = history[:10]

def save_base_path(path_text):
    normalized_path = normalize_input_path(path_text)
    data = load_config_data()
    data["onedrive_path"] = normalized_path
    update_path_history(data, normalized_path)
    save_config_data(data)
    return normalized_path, data

def prompt_for_base_path(existing_path=""):
    while True:
        print("\n" + "="*42)
        print(" 🔧 최상위 폴더 경로 설정")
        print("="*42)
        if existing_path:
            print(f"현재 저장된 경로: {existing_path}")

        path_input = input("\n▶ 원드라이브 최상위 폴더 경로 입력: ").strip()
        normalized_path = normalize_input_path(path_input)

        if is_valid_base_path(normalized_path):
            return save_base_path(normalized_path)

        print("❌ 존재하지 않는 폴더 경로입니다. 다시 입력해주세요.")
        time.sleep(1)

def load_or_request_config(force_change=False):
    data = load_config_data()
    saved_path = normalize_input_path(data.get("onedrive_path", ""))

    if force_change:
        return prompt_for_base_path(saved_path)

    if is_valid_base_path(saved_path):
        update_path_history(data, saved_path)
        save_config_data(data)
        return saved_path, data

    return prompt_for_base_path(saved_path)

def create_onedrive_direct_download(onedrive_link):
    data_bytes64 = base64.b64encode(bytes(onedrive_link, 'utf-8'))
    data_bytes64_String = data_bytes64.decode('utf-8').replace('/','_').replace('+','-').rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_String}/root/content"

def prompt_menu_input(max_index):
    prompt = f"▶ 번호 선택 (0, 9, 1~{max_index}) [ESC:새로고침, q:종료]: "
    if msvcrt is None:
        return input(prompt).strip().lower()

    sys.stdout.write(prompt)
    sys.stdout.flush()
    buffer = []

    while True:
        key = msvcrt.getwch()

        if key in ('\r', '\n'):
            print()
            return ''.join(buffer).strip().lower()

        if key == '\x1b':
            print()
            return 'esc'

        if key == '\x08':
            if buffer:
                buffer.pop()
                sys.stdout.write('\b \b')
                sys.stdout.flush()
            continue

        if key in ('\x00', '\xe0'):
            msvcrt.getwch()
            continue

        buffer.append(key)
        sys.stdout.write(key)
        sys.stdout.flush()

def render_scan_progress(phase_label, percent, current_count, total_count, found_files, start_time, color):
    bar_width = 24
    filled = int(bar_width * percent / 100)
    bar = "█" * filled + "░" * (bar_width - filled)
    elapsed = time.time() - start_time
    sys.stdout.write(
        f"\r{color}🔎 {phase_label} [{bar}] {percent:3d}%{COLOR_RESET} "
        f"폴더 {current_count}/{total_count} / 대상 파일 {found_files}개 / {elapsed:.1f}s"
    )
    sys.stdout.flush()

def emit_scan_progress(progress_callback, phase_label, percent, current_count, total_count, found_files, start_time, color):
    if progress_callback:
        progress_callback({
            "phase_label": phase_label,
            "percent": percent,
            "current_count": current_count,
            "total_count": total_count,
            "found_files": found_files,
            "elapsed": time.time() - start_time,
            "color": color,
        })
    else:
        render_scan_progress(phase_label, percent, current_count, total_count, found_files, start_time, color)

def scan_target_files(base_path, progress_callback=None):
    win_path = base_path
    if os.name == 'nt' and not win_path.startswith("\\\\?\\"):
        win_path = "\\\\?\\" + os.path.abspath(win_path)

    total_dirs = 0
    count_start = time.time()
    last_update = 0
    for _, dirs, files in os.walk(win_path):
        total_dirs += 1
        now = time.time()
        if now - last_update >= 0.05:
            count_percent = min(50, total_dirs)
            emit_scan_progress(progress_callback, "폴더 수 계산 중", count_percent, total_dirs, max(total_dirs, 1), 0, count_start, COLOR_CYAN)
            last_update = now

    total_dirs = max(total_dirs, 1)

    found_files = []
    scanned_dirs = 0
    start_time = count_start
    last_update = 0

    for root, dirs, files in os.walk(win_path):
        scanned_dirs += 1
        now = time.time()
        if now - last_update >= 0.05:
            scan_percent = 50 + int((scanned_dirs / total_dirs) * 50)
            emit_scan_progress(progress_callback, "파일 목록 스캔 중", min(scan_percent, 100), scanned_dirs, total_dirs, len(found_files), start_time, COLOR_GREEN)
            last_update = now

        for file in files:
            ext = splitext(file)[1].lower()
            if ext in TARGET_EXTENSIONS:
                full_path = join(root, file)
                try:
                    found_files.append((full_path, getmtime(full_path)))
                except:
                    continue

    emit_scan_progress(progress_callback, "파일 목록 스캔 중", 100, total_dirs, total_dirs, len(found_files), start_time, COLOR_YELLOW)
    if not progress_callback:
        print()
    found_files.sort(key=lambda x: x[1], reverse=True)
    return found_files[:5]

def format_recent_file_entry(full_path, modified_time, base_path=""):
    clean_full_path = sanitize_display_path(full_path)
    clean_base_path = sanitize_display_path(base_path)

    try:
        relative_dir = os.path.dirname(os.path.relpath(clean_full_path, clean_base_path)) if clean_base_path else os.path.dirname(clean_full_path)
    except ValueError:
        relative_dir = os.path.dirname(clean_full_path)

    if relative_dir in (".", ""):
        relative_dir = "(루트)"

    return {
        "path": clean_full_path,
        "directory": relative_dir.replace("/", "\\"),
        "filename": os.path.basename(clean_full_path),
        "timestamp": time.strftime("%m-%d %H:%M", time.localtime(modified_time)),
        "extension": os.path.splitext(clean_full_path)[1].lower(),
    }

# ==========================================
# 🛠️ 2. 핵심 워커 함수
# ==========================================
def get_connected_devices():
    result = subprocess.run("adb devices", shell=True, capture_output=True, text=True)
    devices = []
    for line in result.stdout.strip().split('\n')[1:]:
        if '\tdevice' in line:
            devices.append(line.split('\t')[0])
    return devices

def get_device_labels(devices=None):
    devices = devices or get_connected_devices()
    return {device: get_device_display_name(device) for device in devices}

def get_device_prop(device, prop_name):
    result = subprocess.run(
        f'adb -s {device} shell getprop {prop_name}',
        shell=True,
        capture_output=True,
        text=True,
        errors='replace'
    )
    return result.stdout.strip()

def get_device_display_name(device):
    model = get_device_prop(device, "ro.product.model") or "UnknownModel"
    os_version = get_device_prop(device, "ro.build.version.release") or "UnknownOS"
    abi64 = get_device_prop(device, "ro.product.cpu.abilist64")
    abi = get_device_prop(device, "ro.product.cpu.abi")

    if abi64:
        arch = "64비트"
    elif "64" in abi:
        arch = "64비트"
    else:
        arch = "32비트"

    model = model.replace(" ", "-")
    return f"{device}_{model}_{os_version}OS_{arch}"

def extract_command_output(result):
    stdout = (result.stdout or "").strip()
    stderr = (result.stderr or "").strip()
    return "\n".join(part for part in [stdout, stderr] if part).strip()

def summarize_failure_reason(output, returncode):
    text = (output or "").strip()
    lower = text.lower()

    reason_map = [
        ("install_failed_no_matching_abis", "32비트 기기/ABI 불일치"),
        ("device offline", "adb 연결 끊김(device offline)"),
        ("device not found", "adb 연결 끊김(device not found)"),
        ("no devices/emulators found", "adb 기기 미연결"),
        ("more than one device/emulator", "복수 기기 충돌"),
        ("install_failed_insufficient_storage", "저장공간 부족"),
        ("insufficient storage", "저장공간 부족"),
        ("broken pipe", "adb 연결 끊김"),
        ("connection reset", "adb 연결 끊김"),
        ("closed", "adb 연결 끊김"),
        ("unauthorized", "adb 디바이스 인증 필요"),
    ]

    for needle, message in reason_map:
        if needle in lower:
            return message

    if text:
        last_line = text.splitlines()[-1].strip()
        return last_line[:80]

    if returncode not in (0, None):
        return f"명령 종료 코드 {returncode}"

    return "원인 미확인"

def set_device_progress(progress_state, lock, display_name, percent, message):
    with lock:
        progress_state[display_name] = {
            "percent": percent,
            "message": message,
        }

def update_device_progress(progress_state, lock, display_name, percent, message, progress_callback=None):
    set_device_progress(progress_state, lock, display_name, percent, message)
    if progress_callback:
        progress_callback(display_name, percent, message)

def render_device_progress(progress_state, device_order):
    lines = []
    for display_name in device_order:
        info = progress_state.get(display_name, {"percent": 0, "message": "대기중"})
        percent = max(0, min(100, int(info["percent"])))
        percent_text = f"{percent:3d}%"
        if "❌" in info["message"]:
            percent_text = f"{COLOR_RED}{percent_text}{COLOR_RESET}"
        lines.append(f"- [{display_name}] {percent_text} {info['message']}")
    return lines

def print_device_progress(progress_state, device_order, previously_rendered):
    lines = render_device_progress(progress_state, device_order)

    if previously_rendered:
        sys.stdout.write(f"\033[{previously_rendered}F")

    for line in lines:
        sys.stdout.write("\033[2K" + line + "\n")

    sys.stdout.flush()
    return len(lines)

def process_device_task(args):
    file_path, device, file_ext, target_name, display_name, progress_state, lock, progress_callback = args
    start_time = time.time()

    try:
        if file_ext == '.apk':
            update_device_progress(progress_state, lock, display_name, 10, "설치 준비중", progress_callback)
            cmd = f'adb -s {device} install -r -d -g "{file_path}"'
            update_device_progress(progress_state, lock, display_name, 55, "APK 설치중", progress_callback)
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, errors='replace')
            output = extract_command_output(result)
            elapsed = time.time() - start_time

            if result.returncode == 0 and "Success" in (result.stdout or ""):
                status = f"✅ 설치 성공 ({elapsed:.2f}s)"
            else:
                reason = summarize_failure_reason(output, result.returncode)
                status = f"❌ 설치 실패 ({reason}, {elapsed:.2f}s)"

            update_device_progress(progress_state, lock, display_name, 100, status, progress_callback)
        elif file_ext == '.obb':
            update_device_progress(progress_state, lock, display_name, 15, "OBB 경로 준비중", progress_callback)
            target_dir = f"/sdcard/Android/obb/{target_name}"
            mkdir_result = subprocess.run(
                f"adb -s {device} shell mkdir -p {target_dir}",
                shell=True,
                capture_output=True,
                text=True,
                errors='replace'
            )
            mkdir_output = extract_command_output(mkdir_result)
            if mkdir_result.returncode != 0:
                elapsed = time.time() - start_time
                reason = summarize_failure_reason(mkdir_output, mkdir_result.returncode)
                status = f"❌ 복사 실패 ({reason}, {elapsed:.2f}s)"
                update_device_progress(progress_state, lock, display_name, 100, status, progress_callback)
                return display_name, status, elapsed

            cmd = f'adb -s {device} push "{file_path}" "{target_dir}/"'
            update_device_progress(progress_state, lock, display_name, 60, "OBB 복사중", progress_callback)
            res = subprocess.run(cmd, shell=True, capture_output=True, text=True, errors='replace')
            output = extract_command_output(res)
            elapsed = time.time() - start_time

            if res.returncode == 0:
                status = f"✅ 복사 성공 ({elapsed:.2f}s)"
            else:
                reason = summarize_failure_reason(output, res.returncode)
                status = f"❌ 복사 실패 ({reason}, {elapsed:.2f}s)"

            update_device_progress(progress_state, lock, display_name, 100, status, progress_callback)
        else:
            elapsed = time.time() - start_time
            status = f"❌ 지원하지 않는 파일 형식 ({elapsed:.2f}s)"
            update_device_progress(progress_state, lock, display_name, 100, status, progress_callback)
    except Exception as e:
        elapsed = time.time() - start_time
        status = f"❌ 오류 ({str(e)[:60]}, {elapsed:.2f}s)"
        update_device_progress(progress_state, lock, display_name, 100, status, progress_callback)

    return display_name, status, time.time() - start_time

def resolve_external_install_input(raw_input):
    sanitized_input = raw_input.strip().replace('"', '')
    if os.path.exists(sanitized_input) and os.path.isfile(sanitized_input):
        return sanitized_input, splitext(sanitized_input)[1].lower(), None
    if sanitized_input.startswith("http"):
        try:
            direct_url = create_onedrive_direct_download(sanitized_input)
            save_path = os.path.join(TEMP_DOWNLOAD_DIR, "temp_download.apk")
            response = requests.get(direct_url, allow_redirects=True)
            with open(save_path, 'wb') as f:
                f.write(response.content)
            return save_path, ".apk", None
        except Exception as e:
            return None, None, f"실패: {e}"
    return None, None, "경로 오류"

def install_to_devices(sel_file, ext, devices=None, progress_callback=None):
    if ext == '.bat':
        subprocess.Popen(f'cmd /c "{sel_file}"', cwd=os.path.dirname(sel_file), shell=True)
        return {
            "success": True,
            "mode": "bat",
            "device_labels": [],
            "results": [],
            "summary": "✅ BAT 실행 완료",
        }

    devices = devices or get_connected_devices()
    if not devices:
        return {
            "success": False,
            "mode": "device",
            "device_labels": [],
            "results": [],
            "summary": "❌ 연결된 기기 없음",
        }

    device_labels = get_device_labels(devices)
    device_order = [device_labels[d] for d in devices]
    progress_state = {
        display_name: {"percent": 0, "message": "대기중"}
        for display_name in device_order
    }
    progress_lock = Lock()

    for display_name in device_order:
        if progress_callback:
            progress_callback(display_name, 0, "대기중")

    args = [
        (sel_file, d, ext, basename(sel_file).replace('.obb',''), device_labels[d], progress_state, progress_lock, progress_callback)
        for d in devices
    ]

    with ThreadPoolExecutor(max_workers=len(args) or 1) as executor:
        futures = [executor.submit(process_device_task, arg) for arg in args]
        results = [future.result() for future in futures]

    return {
        "success": True,
        "mode": "device",
        "device_labels": device_order,
        "results": results,
        "summary": f"✅ {len(device_order)}대 작업 완료",
        "progress_state": progress_state,
    }

def run_selected_install(sel_file, ext):
    if ext == '.bat':
        result = install_to_devices(sel_file, ext)
        print(result["summary"])
        return result["success"]

    devices = get_connected_devices()
    if not devices:
        print("❌ 연결된 기기 없음")
        return False

    device_labels = get_device_labels(devices)
    print(f"\n▶ 실행 대상 기기 {len(devices)}대")
    for label in device_labels.values():
        print(f"   - {label}")

    print(f"\n🚀 {len(devices)}대 병렬 작업 시작...")
    rendered_lines = 0
    progress_state = {
        display_name: {"percent": 0, "message": "대기중"}
        for display_name in [device_labels[d] for d in devices]
    }

    def console_progress_callback(display_name, percent, message):
        nonlocal rendered_lines
        progress_state[display_name] = {"percent": percent, "message": message}
        rendered_lines = print_device_progress(progress_state, list(progress_state.keys()), rendered_lines)

    install_to_devices(sel_file, ext, devices=devices, progress_callback=console_progress_callback)
    print()
    return True

# ==========================================
# 🚀 3. 메인 인터페이스 (무한 루프 적용)
# ==========================================
def main():
    base_path, config_data = load_or_request_config()
    top_5 = scan_target_files(base_path)

    while True: # 💡 무한 루프 시작
        os.system('cls' if os.name == 'nt' else 'clear') 
        print("="*42)
        print(" 🚀 KRAFTON QA 자동 설치기 v8.0 (연속 실행)")
        print("="*42)
        print(f"현재 연결된 최상위 경로: {base_path}")
        path_history = config_data.get("path_history", [])
        if path_history:
            print("최근 저장 경로(최대2건)")
            for entry in path_history[:2]:
                print(f"- {entry.get('path', base_path)}")
        print()

        if not is_valid_base_path(base_path):
            print("❌ 저장된 최상위 경로를 찾을 수 없습니다.")
            base_path, config_data = load_or_request_config(force_change=True)
            top_5 = scan_target_files(base_path)
            continue

        print("[메뉴]")
        print("0.  최상위 폴더 경로 변경")
        print("-" * 42)
        print("최근 apk / obb / bat 파일 TOP 5")
        print("-" * 42)
        if not top_5:
            print("표시할 apk / obb / bat 파일이 없습니다.")
        for i, (f_path, f_time) in enumerate(top_5):
            clean_p = f_path.replace("\\\\?\\", "").replace(base_path, "").lstrip("\\")
            t_str = time.strftime("%m-%d %H:%M", time.localtime(f_time))
            print(f"{i+1}. [{os.path.dirname(clean_p)}]\n   └ {os.path.basename(clean_p)} ({t_str})")
        print("-" * 42)
        print("9.  외부 설치 경로 직접 입력")
        print("esc. 화면 새로고침")
        print("-" * 42)

        u_in = prompt_menu_input(len(top_5))
        if u_in in ['q', 'exit', 'quit']: break
        if u_in == 'esc':
            top_5 = scan_target_files(base_path)
            continue
        
        try:
            if not (u_in in ['0', '9'] or (u_in.isdigit() and 1 <= int(u_in) <= len(top_5))):
                print("⚠️ 잘못된 입력입니다."); time.sleep(1); continue
            choice = int(u_in)
        except: continue

        if choice == 0:
            base_path, config_data = load_or_request_config(force_change=True)
            top_5 = scan_target_files(base_path)
            print("\n✅ 최상위 경로가 변경되었습니다.")
            time.sleep(1)
            continue

        sel_file = ""
        ext = ""

        if choice == 9:
            raw_input = input("\n▶ 링크(https) 또는 파일 경로(C:\\)를 입력: ").strip().replace('"', '')
            sel_file, ext, error_message = resolve_external_install_input(raw_input)
            if error_message:
                print(f"❌ {error_message}")
                time.sleep(2 if "실패" in error_message else 1)
                continue
        else:
            sel_file, _ = top_5[choice-1]
            ext = splitext(sel_file)[1].lower()

        while True:
            run_selected_install(sel_file, ext)
            retry_input = input("\n✅ 모든 작업 완료! 엔터: 메뉴로 돌아가기 / R: 동일 파일 재설치 > ").strip().lower()
            if retry_input == 'r':
                continue
            break

    print("\n👋 프로그램을 종료합니다.")

if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    main()
