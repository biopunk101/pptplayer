# pptplayer

Windows 週會簡報自動連播工具。依照預設的講者順序，自動開啟並播放對應的 PowerPoint 簡報，每份簡報結束後自動繼續下一位。

## 需求

- Windows 10 / 11
- Python 3.9+
- Microsoft PowerPoint（已安裝於本機）
- 相依套件：

```bash
pip install pywin32
```

## 使用方式

### 直接執行 Python

```bash
python play_meeting.py
```

### 使用執行檔（免安裝 Python）

直接執行 `play_meeting.exe`，打包方式請參考 [build py exec.md](build%20py%20exec.md)。

## 操作流程

1. 選擇簡報資料夾
2. 選擇播放順序檔（`.txt`）
3. 預覽比對結果，確認後開始連播
4. 講者關閉放映視窗後，自動播放下一份
5. 全部播放完畢後顯示播放統計（總耗時 / 平均每份）

```
選擇資料夾 → 選擇順序檔 → 預覽比對結果 → 確認開始 → 自動逐一播放
```

## 播放順序檔格式

`sequence.txt` 每行填入一位講者名稱，程式會比對資料夾中**檔名開頭**符合該名稱的 `.ppt` / `.pptx` 檔案。

```
# 以 # 開頭的行為註解，會自動略過
Leo
Jun
Shinra
```

> 比對不區分大小寫，例如 `shinra` 可匹配 `Shinra_20260323.pptx`。

## 設定檔

程式會將上次選擇的資料夾與順序檔路徑儲存於 `~/.pptplayer_config.json`，下次開啟時自動帶入。

## 檔案結構

```
pptplayer/
├── play_meeting.py     # 主程式
├── sequence.txt        # 播放順序範例
├── build py exec.md    # 打包為 exe 的說明
└── README.md
```
