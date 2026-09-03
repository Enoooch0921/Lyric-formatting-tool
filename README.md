# Lyric-formatting-tool
在製作 EasySlides 的歌詞中，適用於製作單行字幕、刪除標點符號、自動空行，以及刪除歌詞結尾的空白。

在第一個文字框中貼入歌詞並點選「一鍵格式化」，結果會出現在第二個文字框。

## 功能

- 自訂行寬並自動排版。
- 拖入 PPTX 後自動辨識多首歌曲、合併重複歌詞投影片，並選擇要匯入的歌曲。
- 辨識 `V1`～`V8`、`Verse 1`～`Verse 8`、`[1]`～`[8]`、`Chorus`、`Bridge`、`Pre Chorus` 等段落標籤。
- 將常見錯字 `Pro Chorus` 修正為 `[prechorus]`。
- 保守地移除明確的中文或英文翻譯行；中英混合與無法判定的內容會保留。
- 可移除每個段落或文字區塊中的雙數內容行，並可復原最近的操作。

線上使用：<https://enoooch0921.github.io/Lyric-formatting-tool/>

### PPTX 匯入

可拖入最多 50 MB 的 `.pptx` 檔案。工具會在瀏覽器本機讀取投影片文字，透過頁碼重置、單行歌名頁與背景變化辨識歌曲，再將完整重複或局部重複的歌詞整理成段落。檔案不會上傳到伺服器。

PPTX 解壓縮使用 [fflate](https://github.com/101arrowz/fflate) 0.8.2（MIT License），程式已隨網站一併提供，不依賴外部 CDN。

## 測試

需要 Node.js 20 或更新版本：

```sh
npm test
```

每次 push 與 pull request 都會透過 GitHub Actions 執行相同測試。
