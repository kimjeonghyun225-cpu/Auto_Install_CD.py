import os
import time
import subprocess
import json
import sys
import requests
import base64
from multiprocessing import Pool
from os.path import basename, join, isdir, abspath, getmtime, splitext

# ==========================================
# ⚙️ 1. 설정 및 경로 관리
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config_qa_installer.json")
TARGET_EXTENSIONS = ('.apk', '.bat', '.obb')
TEMP_DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "Downloads")

if not os.path.exists(TEMP_DOWNLOAD_DIR):
    os.makedirs(TEMP_DOWNLOAD_DIR)

def load_or_request_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data['onedrive_path']
    else:
        print("\n" + "="*42)
        print(" 🔧 [최초 실행] 환경 설정이 필요합니다.")
        print("="*42)
        path_input = input("\n▶ 원드라이브 최상위 폴더 경로 입력: ").strip().replace('"', '').replace("'", "")
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'onedrive_path': path_input}, f, ensure_ascii=False, indent=4)
        return path_input

def create_onedrive_direct_download(onedrive_link):
    data_bytes64 = base64.b64encode(bytes(onedrive_link, 'utf-8'))
    data_bytes64_String = data_bytes64.decode('utf-8').replace('/','_').replace('+','-').rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_String}/root/content"

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
    return f"{device}_{model}_{os_version}버전_{arch}"

def check_apk_installed(package_name, device):
    command = f"adb -s {device} shell pm list packages {package_name}"
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    return package_name in result.stdout

def process_uninstall_task(args):
    package_name, device, display_name = args
    start_time = time.time()
    if check_apk_installed(package_name, device):
        print(f"   └  [{display_name}] '{package_name}' 삭제 중...")
        subprocess.run(f"adb -s {device} uninstall {package_name}", shell=True, capture_output=True)
        status = "✅ 삭제 완료"
    else:
        status = "⏩ 대상 없음"
    return display_name, status, time.time() - start_time

def process_device_task(args):
    file_path, device, file_ext, target_name, display_name = args
    start_time = time.time()
    try:
        if file_ext == '.apk':
            cmd = f'adb -s {device} install -r -d -g "{file_path}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, errors='replace')
            status = "✅ 설치 성공" if "Success" in result.stdout else f"❌ 실패 ({result.stdout[-20:].strip()})"
        elif file_ext == '.obb':
            target_dir = f"/sdcard/Android/obb/{target_name}"
            subprocess.run(f"adb -s {device} shell mkdir -p {target_dir}", shell=True)
            cmd = f'adb -s {device} push "{file_path}" "{target_dir}/"'
            res = subprocess.run(cmd, shell=True, capture_output=True)
            status = "✅ 복사 성공" if res.returncode == 0 else "❌ 복사 실패"
    except: status = "❌ 오류"
    return display_name, status, time.time() - start_time

# ==========================================
# 🚀 3. 메인 인터페이스 (무한 루프 적용)
# ==========================================
def main():
    base_path = load_or_request_config()

    while True: # 💡 무한 루프 시작
        os.system('cls' if os.name == 'nt' else 'clear') 
        print("="*42)
        print(" 🚀 KRAFTON QA 자동 설치기 v8.0 (연속 실행)")
        print("="*42)
        
        win_path = base_path
        if os.name == 'nt' and not win_path.startswith("\\\\?\\"):
            win_path = "\\\\?\\" + os.path.abspath(win_path)

        # 파일 스캔
        found_files = []
        for root, dirs, files in os.walk(win_path):
            for file in files:
                ext = splitext(file)[1].lower()
                if ext in TARGET_EXTENSIONS:
                    full_path = join(root, file)
                    try: found_files.append((full_path, getmtime(full_path)))
                    except: continue

        found_files.sort(key=lambda x: x[1], reverse=True)
        top_5 = found_files[:5]
        
        print("[메뉴] (종료: q)")
        print("0.  앱 수동 삭제 (패키지명)")
        print("9.  외부 링크 설치 OR 로컬 경로 직접 입력")
        print("-" * 42)
        for i, (f_path, f_time) in enumerate(top_5):
            clean_p = f_path.replace("\\\\?\\", "").replace(base_path, "").lstrip("\\")
            t_str = time.strftime("%m-%d %H:%M", time.localtime(f_time))
            print(f"{i+1}. [{os.path.dirname(clean_p)}]\n   └ {os.path.basename(clean_p)} ({t_str})")
        print("-" * 42)

        u_in = input(f"▶ 번호 선택 (0, 9, 1~{len(top_5)}) [q:종료]: ").strip().lower()
        if u_in in ['q', 'exit', 'quit']: break
        
        try:
            if not (u_in in ['0', '9'] or (u_in.isdigit() and 1 <= int(u_in) <= len(top_5))):
                print("⚠️ 잘못된 입력입니다."); time.sleep(1); continue
            choice = int(u_in)
        except: continue

        devs = get_connected_devices()
        if not devs: 
            print("❌ 연결된 기기 없음"); input("\n엔터를 누르면 메뉴로 돌아갑니다."); continue
        device_labels = {device: get_device_display_name(device) for device in devs}
        print(f"\n▶ 실행 대상 기기 {len(devs)}대")
        for label in device_labels.values():
            print(f"   - {label}")

        sel_file = ""
        ext = ""

        if choice == 9:
            raw_input = input("\n▶ 링크(https) 또는 파일 경로(C:\\)를 입력: ").strip().replace('"', '')
            if os.path.exists(raw_input) and os.path.isfile(raw_input):
                sel_file, ext = raw_input, splitext(raw_input)[1].lower()
            elif raw_input.startswith("http"):
                try:
                    direct_url = create_onedrive_direct_download(raw_input)
                    save_path = os.path.join(TEMP_DOWNLOAD_DIR, "temp_download.apk")
                    print("⏳ 다운로드 중..."); r = requests.get(direct_url, allow_redirects=True)
                    with open(save_path, 'wb') as f: f.write(r.content)
                    sel_file, ext = save_path, ".apk"
                except Exception as e: print(f"❌ 실패: {e}"); time.sleep(2); continue
            else: print("❌ 경로 오류"); time.sleep(1); continue
        
        elif choice == 0:
            pkg = input("\n▶ 삭제할 패키지명: ").strip()
            if pkg:
                with Pool() as p:
                    res = p.map(process_uninstall_task, [(pkg, d, device_labels[d]) for d in devs])
                for d, s, t in res: print(f"- [{d}] {s} ({t:.2f}s)")
                input("\n작업 완료! 엔터를 누르면 메뉴로 돌아갑니다."); continue
        else:
            sel_file, _ = top_5[choice-1]
            ext = splitext(sel_file)[1].lower()

        # 설치 실행
        if ext == '.bat':
            subprocess.Popen(f'cmd /c "{sel_file}"', cwd=os.path.dirname(sel_file), shell=True)
            print("✅ BAT 실행 완료")
        else:
            print(f"\n🚀 {len(devs)}대 병렬 작업 시작...")
            args = [(sel_file, d, ext, basename(sel_file).replace('.obb',''), device_labels[d]) for d in devs]
            with Pool() as p:
                res = p.map(process_device_task, args)
            for d, s, t in res: print(f"- [{d}] {s} ({t:.2f}s)")

        input("\n✅ 모든 작업 완료! 엔터를 누르면 메뉴로 돌아갑니다.")

    print("\n👋 프로그램을 종료합니다.")

if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    main()