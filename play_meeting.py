import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import json

DEFAULT_FOLDER = "C:\\Users\\ShinraLuo\\OneDrive"
DEFAULT_SEQUENCE = "C:\\Users\\ShinraLuo\\pptplayer\\sequence.txt"
CONFIG_FILE = os.path.expanduser("~/.pptplayer_config.json")


# ── Dialog 工具 ──────────────────────────────────────────────────────────────

def choose_folder(default):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="請選擇簡報資料夾", initialdir=default)
    root.destroy()
    return folder if folder else None


def choose_file(default):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file = filedialog.askopenfilename(
        title="請選擇播放順序檔",
        initialdir=os.path.dirname(default),
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    root.destroy()
    return file if file else None


def confirm_dialog(title, message):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askokcancel(title, message)
    root.destroy()
    return result


def notify(message):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("週會簡報連播", message)
    root.destroy()


def alert(message):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror("週會簡報連播", message)
    root.destroy()


# ── Config 工具 ──────────────────────────────────────────────────────────────

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"last_folder": DEFAULT_FOLDER, "last_seq": DEFAULT_SEQUENCE}

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


# ── 核心邏輯 ───────────────────────────────────────────────────────────────────

def load_speakers(seq_file):
    with open(seq_file, "r", encoding="utf-8") as f:
        return [l.strip() for l in f if l.strip() and not l.strip().startswith("#")]


def match_files(speakers, folder):
    ppt_files = [f for f in os.listdir(folder) if f.endswith(('.ppt', '.pptx'))]
    return [(s, next((f for f in ppt_files if f.lower().startswith(s.lower())), None))
            for s in speakers]


def play_file(ppt_app, file_path):
    import pywintypes

    presentation = None
    # 加入重試機制來開啟檔案，避免前一份剛關閉時 PowerPoint 仍在忙碌
    for _ in range(20):
        try:
            presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), WithWindow=True)
            break
        except pywintypes.com_error:
            time.sleep(0.5)
            
    if not presentation:
        return True  # 如果開啟失敗則跳過，繼續播放下一份

    # 加上重試機制，避免 PowerPoint 忙碌中導致「接收者已拒絕這個呼叫」錯誤
    for _ in range(20):  # 最多嘗試 10 秒
        try:
            presentation.SlideShowSettings.Run()
            break
        except pywintypes.com_error:
            time.sleep(0.5)
            
    while True:
        try:
            if ppt_app.SlideShowWindows.Count == 0:
                break
        except Exception:
            # 發生例外 (例如使用者直接把整個 PowerPoint 關閉了)，回傳 False 代表中斷全部連播
            return False
        time.sleep(0.5)

    try:
        presentation.Saved = True
        presentation.Close()
    except Exception:
        pass
        
    return True


# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main():
    config = load_config()

    # 1. 選擇資料夾
    folder = choose_folder(config.get("last_folder", DEFAULT_FOLDER))
    if not folder:
        return
    config["last_folder"] = folder

    # 2. 選擇順序檔
    seq_file = choose_file(config.get("last_seq", DEFAULT_SEQUENCE))
    if not seq_file:
        return
    config["last_seq"] = seq_file

    save_config(config)

    # 3. 比對檔案
    try:
        speakers = load_speakers(seq_file)
        matches = match_files(speakers, folder)
    except Exception as e:
        alert(str(e))
        return

    # 4. 預覽並確認
    lines = []
    for i, (speaker, fname) in enumerate(matches, 1):
        lines.append(f"{i}. {fname}" if fname else f"{i}. ⚠ 找不到 [{speaker}]")

    found = sum(1 for _, f in matches if f)
    preview = "\n".join(lines)
    summary = f"找到 {found}/{len(matches)} 份簡報，確認開始連播？\n\n{preview}"

    if not confirm_dialog("週會簡報連播", summary):
        return

    import win32com.client
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
    except Exception as e:
        alert(f"無法啟動 PowerPoint: {e}")
        return

    # 5. 依序播放
    start_time = time.time()
    played_count = 0
    is_aborted = False

    for _, fname in matches:
        if fname:
            continue_play = play_file(ppt_app, os.path.join(folder, fname))
            if not continue_play:
                is_aborted = True
                break
            played_count += 1

    end_time = time.time()
    total_seconds = end_time - start_time

    stats_msg = ""
    if played_count > 0:
        total_m, total_s = divmod(total_seconds, 60)
        avg_m, avg_s = divmod(total_seconds / played_count, 60)
        stats_msg = (f"\n\n📊 播放統計："
                     f"\n總計耗時：{int(total_m)} 分 {int(total_s)} 秒"
                     f"\n平均每份：{int(avg_m)} 分 {int(avg_s)} 秒"
                     f"\n實際播放：{played_count} 份")

    if is_aborted:
        notify(f"已手動中斷簡報連播。{stats_msg}")
    else:
        notify(f"所有簡報播放完畢！{stats_msg}")

    try:
        if ppt_app.Presentations.Count == 0:
            ppt_app.Quit()
    except Exception:
        pass


if __name__ == "__main__":
    main()
