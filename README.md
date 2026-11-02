# 🎲 AI-TRPG-CoAgent-System

一個基於大型語言模型（LLM）的多代理人（Multi-Agent）《龍與地下城》（D&D 5e）跑團系統。

---

## 🎯 專題簡介
解決 TRPG（桌上角色扮演遊戲）玩家「時間難對齊、人數湊不齊」的痛點。本系統利用 AI 劃分兩大角色：
1. **DM（地下城主）Agent**：負責世界觀描繪、NPC 扮演、環境判定與戰鬥引導。
2. **Player Subagent（AI 隊友）**：擁有獨立角色卡與內心思考視窗，在不被 DM 劇透的情況下與玩家一同冒險與決策。

---

## 🛠️ 系統架構與 Prompt 設計
- **主視窗（DM 視窗）**：管理遊戲狀態、HP/AC 數值判定、環境敘事。
- **子視窗（AI 隊友 Agent）**：注入角色性格、背景與限制視角 Prompt，即時回應玩家互動。

---

## 🚀 快速開始
1. 複製 Prompt 模板至 LLM 環境。
2. 選擇預設劇本（如《失落的哥布林礦坑》）。
3. 召喚 AI 隊友並開啟冒險！

---

## 🌐 線上互動網站 (GitHub Pages)
👉 **點此進入互動網站：https://411383049-source.github.io/AI-TRPG-CoAgent-System/**
*(亦可直接雙擊本地的 [`index.html`](./index.html) 開啟完整互動網頁！)*

網站內建四大功能分頁：
1. **⚔️ 雙視窗互動跑團模擬器**：左視窗為「地下城主 (DM) ＆ 隊伍頻道」，右視窗即時顯示「獨立 AI 隊友【萊拉】的無劇透內心戲與視角」，並附帶 D20 擲骰器與互動戰術按鈕！
2. **📄 專案報告 (5大重點)**：線上閱讀並支援一鍵下載 Word (`.docx`) 與列印 PDF。
3. **📊 11/02 分享投影片**：內嵌完整簡報預覽，支援全螢幕播放與下載 PowerPoint (`.pptx`)。
4. **🛠️ Prompt 與架構設計**：展示 Multi-Agent 記憶隔離架構圖與可複製的 Subagent Prompt 模板。

---

## 📂 專案檔案導覽

| 檔案 / 資料夾 | 說明 |
| :--- | :--- |
| 🌐 [`index.html`](./index.html) | **專案主網站入口**：整合雙視窗 D&D 跑團互動模擬器、報告、投影片與架構圖 |
| 📄 [`report/project_report.docx`](./report/project_report.docx) | **完整專案報告 (Word 檔)**：包含做什麼、給誰用、怎麼跟 AI 協作、查了什麼、下一步規劃 |
| 📄 [`report/project_report.md`](./report/project_report.md) | **完整專案報告 (Markdown 版)**：方便直接在 GitHub 上閱讀 |
| 🌐 [`report/project_report.html`](./report/project_report.html) | **完整專案報告 (可列印網頁版)**：支援直接轉存 PDF |
| 📊 [`slides/presentation_1102.pptx`](./slides/presentation_1102.pptx) | **11/02 分享投影片 (PowerPoint 檔)**：課堂/活動分享專用簡報 |
| 🌐 [`slides/presentation_1102.html`](./slides/presentation_1102.html) | **11/02 互動式網頁投影片**：可直接在瀏覽器全螢幕播放 |
| 💬 [`chat_logs/main_dm_room.md`](./chat_logs/main_dm_room.md) | **主聊天室完整對話紀錄**：玩家與 DM 的互動，以及 AI 隊友的轉播演出 |
| 🧝‍♀️ [`chat_logs/companion_lyra_room.md`](./chat_logs/companion_lyra_room.md) | **AI 隊友【萊拉】獨立聊天室紀錄**：揭密子代理的 System Prompt 與無劇透視角思考過程 |
