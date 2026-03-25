# play_meeting.exe 打包說明

## 環境需求

- Python 3.9+
- PyInstaller 6.x
- pywin32（提供 `win32com`、`pywintypes`）

## 安裝依賴

```bash
pip install pyinstaller pywin32
```

## 打包指令

在 `pptplayer` 目錄下執行：

```bash
python -m PyInstaller --onefile --windowed --name play_meeting \
    --hidden-import win32com \
    --hidden-import win32com.client \
    --hidden-import pywintypes \
    --hidden-import win32api \
    play_meeting.py
```

### 參數說明

| 參數 | 說明 |
|------|------|
| `--onefile` | 打包成單一 `.exe`，無需額外資料夾 |
| `--windowed` | 不顯示黑色 Console 視窗 |
| `--name play_meeting` | 指定輸出檔名 |
| `--hidden-import` | 明確引入 pywin32 相關模組（PyInstaller 無法自動偵測動態載入） |

## 輸出位置

```
pptplayer/
├── dist/
│   └── play_meeting.exe   ← 執行檔在這裡
├── build/                 ← 暫存檔，可刪除
└── play_meeting.spec      ← PyInstaller 設定檔（可重複使用）
```

## 重新打包

若程式碼有修改，重新執行上方指令即可，PyInstaller 會自動覆蓋舊的輸出。

或直接使用已產生的 `.spec` 檔：

```bash
python -m PyInstaller play_meeting.spec
```
