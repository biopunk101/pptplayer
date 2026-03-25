# pptplayer

macOS 週會簡報自動連播工具。依照預設的講者順序，自動開啟並播放對應的 PowerPoint 簡報，每份簡報結束後自動繼續下一位。

## 需求

- macOS
- Python 3（系統內建即可）
- Microsoft PowerPoint for Mac

## 使用方式

```bash
python3 play_meeting.py
```

1. 選擇簡報資料夾
2. 選擇播放順序檔（`.txt`）
3. 確認比對結果後點選「開始連播」

## 播放順序檔格式

`sequence.txt` 每行填入一位講者名稱，程式會比對資料夾中**檔名開頭**符合該名稱的 `.ppt` / `.pptx` 檔案。

```
# 以 # 開頭的行為註解，會自動略過
Leo
Jun
Shinra
```

> 比對不區分大小寫，例如 `shinra` 可匹配 `Shinra_20260323.pptx`。

## 操作流程

```
選擇資料夾 → 選擇順序檔 → 預覽比對結果 → 確認開始 → 自動逐一播放
```

講者可自由使用簡報筆翻頁，關閉放映視窗後自動進入下一份。
