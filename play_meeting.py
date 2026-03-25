import os
import subprocess

DEFAULT_FOLDER = "/Users/shinra/Library/CloudStorage/OneDrive-HCmedInnovationsCo.,Ltd/週會資料-研發部/260323/"
DEFAULT_SEQUENCE = "/Users/shinra/Workbench/python/pptplayer/sequence.txt"


# ── AppleScript 工具 ────────────────────────────────────────────────────────────

def run_as(script):
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return r.stdout.strip(), r.returncode


def choose_folder(default):
    out, code = run_as(f'''
        tell application "Finder" to activate
        try
            POSIX path of (choose folder ¬
                with prompt "請選擇簡報資料夾" ¬
                default location POSIX file "{default}")
        on error
            return ""
        end try
    ''')
    return out.rstrip("/") if code == 0 and out else None


def choose_file(default):
    out, code = run_as(f'''
        tell application "Finder" to activate
        try
            POSIX path of (choose file ¬
                with prompt "請選擇播放順序檔" ¬
                default location POSIX file "{default}")
        on error
            return ""
        end try
    ''')
    return out if code == 0 and out else None


def confirm_dialog(title, message):
    _, code = run_as(f'''
        display dialog "{message}" ¬
            with title "{title}" ¬
            buttons {{"取消", "開始連播"}} ¬
            default button "開始連播"
    ''')
    return code == 0


def notify(message):
    run_as(f'display notification "{message}" with title "週會簡報連播"')


def alert(message):
    run_as(f'display alert "週會簡報連播" message "{message}" as critical buttons {{"確定"}}')


# ── 核心邏輯 ───────────────────────────────────────────────────────────────────

def load_speakers(seq_file):
    with open(seq_file, "r", encoding="utf-8") as f:
        return [l.strip() for l in f if l.strip() and not l.strip().startswith("#")]


def match_files(speakers, folder):
    ppt_files = [f for f in os.listdir(folder) if f.endswith(('.ppt', '.pptx'))]
    return [(s, next((f for f in ppt_files if f.lower().startswith(s.lower())), None))
            for s in speakers]


def play_file(file_path):
    subprocess.run(['osascript', '-e', f'''
        set macPath to POSIX file "{file_path}" as alias
        tell application "Microsoft PowerPoint"
            activate
            open macPath
            delay 2
            set thePresentation to active presentation
            run slide show slide show settings of thePresentation
            repeat while (count of slide show windows) > 0
                delay 0.5
            end repeat
            close thePresentation saving no
        end tell
    '''])


# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main():
    # 1. 選擇資料夾
    folder = choose_folder(DEFAULT_FOLDER)
    if not folder:
        return

    # 2. 選擇順序檔
    seq_file = choose_file(DEFAULT_SEQUENCE)
    if not seq_file:
        return

    # 3. 比對檔案
    try:
        speakers = load_speakers(seq_file)
        matches = match_files(speakers, folder)
    except Exception as e:
        alert(str(e).replace('"', "'"))
        return

    # 4. 預覽並確認
    lines = []
    for i, (speaker, fname) in enumerate(matches, 1):
        lines.append(f"{i}. {fname}" if fname else f"{i}. ⚠ 找不到 [{speaker}]")

    found = sum(1 for _, f in matches if f)
    preview = "\\n".join(lines).replace('"', "'")
    summary = f"找到 {found}/{len(matches)} 份簡報，確認開始連播？\\n\\n{preview}"

    if not confirm_dialog("週會簡報連播", summary):
        return

    # 5. 依序播放
    for _, fname in matches:
        if fname:
            play_file(os.path.join(folder, fname))

    notify("所有簡報播放完畢！")


if __name__ == "__main__":
    main()
