# 見習證書發送工具

2026 資訊月 × 臺灣教育科技展「大專碩博 AI × 學習產業見習」時數證書的批次發送工具。
上傳報名名單與問卷回覆 CSV → 對照紙本簽到退 → 產出 PDF 證書 → 以 Gmail 批次寄出。

**使用網址：https://yasaki1234-source.github.io/itmonth-cert-sender/**（登入帳密請洽 EAT 中心）

## 架構

| 層 | 位置 | 說明 |
|---|---|---|
| 網頁介面 | `docs/index.html`（GitHub Pages） | 以 `fetch` 呼叫後端，不帶 Google 登入資訊，瀏覽器登入任何 Google 帳號都不影響 |
| 後端 API | `apps-script/Code.gs`（Google Apps Script） | `doPost` 接收 `{fn, args}`，處理登入、比對、PDF、寄信；資料存於試算表與雲端硬碟 |
| 證書範本 | `template/gen_template.py` 產生的 pptx | 於網頁「設定」上傳後自動轉成 Google 簡報，`{{姓名}}` 等欄位自動套入 |

零成本：GitHub Pages 與 Google Apps Script 皆免費，不需外部廠商。

## 檔案

| 路徑 | 說明 |
|---|---|
| `apps-script/Code.gs` | 後端程式 |
| `apps-script/Index.html` | 舊網址的「已換新網址」說明頁 |
| `apps-script/appsscript.json` | Apps Script 專案設定 |
| `docs/index.html` | 網頁介面（上傳比對／確認名單／產出寄送／紀錄／設定） |
| `template/gen_template.py` | 證書範本產生程式（python-pptx） |
| `tests/` | 比對邏輯與 API 路由測試：`node tests/test_logic.js`、`node tests/test_dopost.js` |

## 比對規則

| 結果 | 條件 | 核發預設 |
|---|---|---|
| 符合 | 簽到 + 簽退 + 問卷 + Email 正確 | 自動勾選 |
| 待確認 | 部分完成、以姓名比對、未在報名名單、手動新增、缺 Email | 人工判斷 |
| 不符合 | 無任何簽到退紀錄 | 不勾 |

比對順序：編號 → Email → 手機 → 姓名（同名不比對）。重複上傳會保留編號、勾選、手動修改與寄送紀錄；已寄出者不會重複寄送。

## 更新程式

- **網頁介面**：修改 `docs/index.html` 推上 GitHub，約 1 分鐘後生效。
- **後端**：修改 `apps-script/Code.gs` 後貼回 Apps Script 編輯器 → 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（API 網址不變）。
