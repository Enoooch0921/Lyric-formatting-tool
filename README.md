# Lyric-formatting-tool
在製作 EasySlides 的歌詞中，適用於製作單行字幕、刪除標點符號、自動空行，以及刪除歌詞結尾的空白。

在第一個文字框中貼入歌詞並點選「一鍵格式化」，結果會出現在第二個文字框。

## 功能

- 自訂行寬並自動排版。
- 辨識 `V1`、`Verse 12`、`[6]`、`Chorus`、`Bridge`、`Pre Chorus` 等段落標籤。
- 將常見錯字 `Pro Chorus` 修正為 `[prechorus]`。
- 保守地移除明確的中文或英文翻譯行；中英混合與無法判定的內容會保留。
- 可移除每個段落或文字區塊中的雙數內容行，並可復原最近的操作。

線上使用：<https://enoooch0921.github.io/Lyric-formatting-tool/>

## 測試

需要 Node.js 20 或更新版本：

```sh
npm test
```

每次 push 與 pull request 都會透過 GitHub Actions 執行相同測試。
