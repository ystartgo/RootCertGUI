# Root Store Viewer / RootCertGUI

Windows 根憑證庫視覺化 / Baseline / Diff / Mozilla 比對 / 風險分類 / 白名單 / 多語 / 匯出。

一個以 PowerShell + WPF 實作的安全觀測與稽核輔助工具，協助您：
- 建立根憑證 Baseline，追蹤新增 / 移除 / 替換
- 與 Mozilla 公共信任根集合進行差集比對
- 標註過期 / 即將過期 / 長效期憑證
- 快速辨識「白名單 / 公認 CA / 不常見」憑證
- 匯出 JSON / CSV / DER，供調查或稽核紀錄
- 多語即時切換 (zh-TW / zh-CN / en)，支援暗色主題與縮放

> 專案 Repo：<https://github.com/ystartgo/RootCertGUI>
- 由 AI GPT-5 協助撰寫除錯



---

## 專案資訊

| 項目 | 資料 |
|------|------|
| 專案名稱 | RootCertGUI |
| 作者 | startgo |
| Email | [startgo@yia.app](mailto:startgo@yia.app) |
| 授權 | GPLv3 |
| 版本 | 1.4.4 (Patched+Filters+PKI) |
| 建立時間 | 2025-08-20 |
| 最近更新 | 2025-09-05 |
| Repository | https://github.com/ystartgo/RootCertGUI |
| 主執行腳本 | `RootCertGUI.ps1` |
| Mozilla certdata.txt | https://hg-edge.mozilla.org/mozilla-central/raw-file/tip/security/nss/lib/ckfw/builtins/certdata.txt |
---

## 核心功能摘要

| 類別 | 功能 |
|------|------|
| 掃描 | 讀取 Windows LocalMachine\Root 憑證 |
| Baseline | 建立 / 重新載入 / 保存基準 JSON |
| Diff | Added / Removed / Replaced（依指紋 + NotAfter / Issuer 判斷） |
| Mozilla | 多來源下載 (`cacert.pem`, `certdata.txt`) 解析並去重 |
| 差集 | Local≠Mozilla / Mozilla≠Local 雙向獨有清單 |
| 風險分類 | 過期 / 90 天內即將過期 / 長效期 (>15 年) |
| 顏色提示 | 過期紅 / 即將過期黃 / 白名單深綠 / 公認淺綠 / 不常見淡米 |
| 白名單 | 以 SHA256 或 Subject 納入/移除，分類即時變更 |
| 匯出 | JSON / CSV / 多檔 DER（自動命名 & 避免衝突） |
| 多語 | zh-TW / zh-CN / en 即時切換（不需重啟） |
| 外觀 | 字體大小 / ZoomFactor / Light-Dark Theme |
| Enterprise PKI | 快捷開啟 `pkiview.msc`（存在時） |
| 設定保存 | `gui_settings.json` 保留語言 / 主題 / 字體 / 路徑 |
| Wrapper | `Invoke-ScanLocal`, `Invoke-Diff` 等簡易函式 |

---

## 下載與執行

### 1. 取得程式碼
```powershell
git clone https://github.com/ystartgo/RootCertGUI.git
cd RootCertGUI
```

### 2. 執行（建議系統管理員權限）
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
powershell -ExecutionPolicy Bypass -File .\RootCertGUI.ps1
```
<img width="882" height="426" alt="image" src="https://github.com/user-attachments/assets/bee307a7-91ac-4792-9bbf-4d1586b1adf4" />

### 3. 首次啟動自動流程
1. 無 `baseline_roots.json` → 自動掃描建立  
2. 無 `mozilla_roots.json` → 建立 placeholder（可再點線上更新）  
3. 無 `whitelist.json` → 建立空白白名單  
4. 自動一次掃描 + 風險分類  
5. 可進一步按：Local Update、Mozilla Compare、Risk Classify、Diff
<img width="1920" height="1040" alt="image" src="https://github.com/user-attachments/assets/38640f50-b966-4bcd-bd8d-d6116243e1e0" />

---

## 截圖建議列表（請自行補圖）

| 圖片 | 說明 |
|------|------|
| `docs/img/screenshot_main.png` | 主視窗（工具列 / 統計 / 日誌） |
| `docs/img/screenshot_diff.png` | Diff Added/Removed/Replaced |
| `docs/img/screenshot_unique.png` | Local Unique / Mozilla Unique |
| `docs/img/screenshot_risk.png` | 風險分類 |
| `docs/img/screenshot_whitelist.png` | 白名單加入/移除 |
| `docs/img/screenshot_dark.png` | 暗色主題 |


---

## 常見使用情境

| 需求 | 操作流程 |
|------|----------|
| 監控根憑證異常新增 | 定期啟動 → Local Update → 查看 Added / Local Unique |
| 稽核過期與即將過期 | Risk Classify 或風險快速按鈕 |
| 檢視企業自簽或不常見 | Local Unique + 顏色（不常見 / 白名單） |
| 蒐集外部佐證 | 匯出 JSON / CSV / DER 提供證據 |
| 企業 PKI 健康檢查 | Enterprise PKI 按鈕開啟 `pkiview.msc` |
| 標示例外信任 | 加入白名單（不更動 OS 信任狀態） |

---

## 快捷鍵

| 快捷鍵 | 功能 |
|--------|------|
| F5 | Diff（Baseline vs Current） |
| F6 | 風險分類重新計算 |
| F7 | Mozilla 比對（建立雙向差集） |
| F8 | 語言資源診斷輸出到 Log |
| Ctrl + L | 重新載入 Baseline |
| List 雙擊列 | 彈出 JSON 查看 |

---

## 匯出與檔案

| 檔案 | 說明 |
|------|------|
| `baseline_roots.json` | 基準快照 |
| `mozilla_roots.json` | Mozilla 憑證集合 |
| `whitelist.json` | 白名單（allow 陣列） |
| `gui_settings.json` | GUI 偏好設定 |
| `output\export_*.json` | 選擇憑證 JSON 匯出 |
| `output\export_*.csv` | 選擇憑證 CSV 匯出 |
| `output\der_*/*.cer` | 成批 DER 憑證檔 |

---

## 憑證資料模型（精簡欄位）

| 欄位 | 說明 |
|------|------|
| subject / issuer | DN |
| sha256 | DER SHA-256 指紋（主鍵） |
| serial | 序號（移除空白） |
| not_before / not_after | ISO 8601 時間 |
| is_ca / is_root | 是否 CA / 自簽根 |
| eku | Extended Key Usage 陣列 |
| path_length | Basic Constraints PathLen |
| risk_status | 過期 / 即將過期 / 正常 (本地化) |
| ca_category | 白名單 / 公認 / 並不常見 / 其他 |
| base64_raw | Base64 (DER) |

---

## 原始碼段落結構 (對應註解)

| 段落 | 內容 |
|------|------|
| 第1段 | 全域狀態 / 語言資源 / UI-Log |
| 第2段 | Baseline / Mozilla / Whitelist / 設定存取 |
| 第3段 | 掃描 / Diff / 風險 / 統計基礎 |
| 第4段 | 白名單測試 / 公認 CA / 不常見分類 |
| 第5段 | 語言切換與欄位標題 |
| 第6段 | 主視覺 GUI 建構 |
| 第7段 | 憑證列表子視窗 + 篩選 + 匯出 |
| 第8段 | Help / Theme / 快捷鍵 / 啟動流程 |
| 第9段 | Wrapper 相容函式 |

---

## 版本重點

| 版本 | 更新 |
|------|------|
| 1.4.x | 重構 GUI / 自動風險分類 / 多來源 Mozilla / Unique 視窗 |
| 1.4.1 | 自動風險執行 / 色彩調整 / 空集合提示 |
| 1.4.3+ | Enterprise PKI / 快速篩選（白名單/公認/不常見） |
| 1.4.4 | 列表視窗預設最大化、語言資源補強、pkiview 整合 |

---

## 安全與注意事項

| 項目 | 說明 |
|------|------|
| 平台限制 | 僅 Windows（使用 WPF） |
| 權限 | 非系統管理員可能掃描不完整 |
| Mozilla 更新 | 需可連線外部來源 |
| 風險邏輯 | 閾值（90 天 / 15 年）可於程式修改 |
| 白名單 | 只影響顯示分類，不改變系統信任 |
| 離線使用 | 可手動放入 `mozilla_roots.json` |

---

## Roadmap

- [ ] Log 行數上限 + 清除按鈕
- [ ] 進階搜尋（正則 / 多條件）
- [ ] 自訂風險閾值 UI
- [ ] HTML / Markdown 報告輸出
- [ ] 加入其他公共信任來源 (Apple / Microsoft CTL)
- [ ] CRL / OCSP 查驗延伸
- [ ] 無 GUI 命令列模式
- [ ] 多執行緒 / 非同步下載優化

---

## 已知限制

| 限制 | 描述 |
|------|------|
| 非跨平台 | 依賴 WPF 與 .NET Framework |
| 延伸欄位 | 尚未解析全部憑證延伸 (Policies / Name Constraints...) |
| certdata 解析 | 若上游格式改變需調整 Parser |
| 大量憑證 | 非虛擬化極大量 (>萬筆) 尚未壓測需調整 |

---

## 範例操作（在 GUI 之外使用函式）

```powershell
# 重新掃描
$Global:AppState.CurrentScanObj = Scan-LocalRootStore

# 建立或更新 Baseline
Save-Baseline -BaselineObj $Global:AppState.CurrentScanObj

# 計算差異
$diff = Compute-RootStoreDiff -BaselineObj $Global:AppState.BaselineObj -CurrentObj $Global:AppState.CurrentScanObj
$diff.added | ConvertTo-Json -Depth 10 | Out-File -Encoding UTF8 .\output\added.json
```

---

## 國際化新增指引

1. 於 `$Global:LangResources` 三語節點新增鍵值  
2. 介面文字呼叫 `(L 'KeyName')`  
3. 具參數訊息：`UI-Log "" "KeyName" @(arg1,arg2,...)`  
4. 務必同步維護 zh-TW / zh-CN / en 避免缺漏  

---

## FAQ

| 問題 | 解答 |
|------|------|
| 為何顯示未以系統管理員執行？ | 權限不足 → 可能缺少某些憑證視圖 |
| Mozilla Compare 無結果？ | 來源下載失敗或無可解析憑證，請重新 Update |
| 「並不常見」定義？ | 非白名單、非常見公認 CA 名稱模式、或屬於 Local≠Mozilla 差集 |
| Long Valid 想改？ | 修改 `Classify-Risks` 呼叫參數或函式內 LongYears |
| 可離線更新 Mozilla 嗎？ | 可：手動備妥 `mozilla_roots.json` 置於同目錄 |

---

## 授權

本專案以 **GNU GPLv3** 授權釋出。  
詳細條款：<https://www.gnu.org/licenses/gpl-3.0.html>

---

## 免責聲明

本工具僅供安全檢視、教育與系統管理輔助；分類與判讀不保證符合所有合規或稽核標準。  
在實際修改系統憑證前請務必備份並審慎評估。

---

## 貢獻 & 聯絡

- Issue / PR：<https://github.com/ystartgo/RootCertGUI>
- Email：startgo@yia.app  
- 歡迎提交改進、功能建議與語系補強。

若此工具對您有幫助，請幫忙 Star 支持後續開發。

---

（README 結束）

---

### 以 PowerShell 直接建立 (UTF-8 BOM)
```powershell
$readme = @'
(請貼上 README.md 內容)
'@
Set-Content -Path .\README.md -Value $readme -Encoding UTF8
```
