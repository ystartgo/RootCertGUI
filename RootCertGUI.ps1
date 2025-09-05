#第1段落 標題 / 全域狀態 與 語言資源 (含新版 UI-Log / StateFlags) 開始
<#
RootCertGUI.ps1
Version : 1.4.4 (Patched+Filters+PKI)
Author  : startgo (startgo@yia.app)
License : GPLv3
Repo    : https://github.com/ystartgo
Created : 2025-08-20
Modified : 2025-09-05
Description:
  - Windows 根憑證庫視覺化 / Baseline / Diff / Mozilla 比對 / 風險分類 / 白名單
  - 多語 (zh-TW / zh-CN / en)
  - 六排工具列 (Baseline | Whitelist | Mozilla | 操作 | 檢視+外觀 | Output)
  - 憑證風險標註 + 匯出 JSON/CSV/DER
  - Mozilla 線上更新 (多來源：cacert.pem / certdata.txt)
  - 獨有清單 UniqueMode
  - 新增：TLS 1.2/1.3 初始化、PEM/Certdata 解析強化、下載結果 HTML 偵測
  - 1.4.1: 風險檢視自動運行、未比對提示、Row5 重排、顏色調整(公認CA淺綠)、空集合提示
  - 1.4.3+: Enterprise PKI 按鈕 / 檢視快速篩選（白名單/公認/不常見）
#>

Remove-Variable -Name AppState -Scope Global -ErrorAction SilentlyContinue
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- TLS 啟用 ---
try {
    $sp = [Net.ServicePointManager]::SecurityProtocol
    $want = [Net.SecurityProtocolType]::Tls12
    try { $want = $want -bor [Net.SecurityProtocolType]::Tls13 } catch {}
    if (($sp -band $want) -ne $want){
        [Net.ServicePointManager]::SecurityProtocol = $sp -bor $want
        Write-Host "[TLS] SecurityProtocol -> $([Net.ServicePointManager]::SecurityProtocol)"
    }
} catch {
    Write-Host "[TLS] 設定失敗: $($_.Exception.Message)"
}

$Global:AppState = [ordered]@{
    Version        = "1.4.4"
    BaselinePath   = ".\baseline_roots.json"
    WhitelistPath  = ".\whitelist.json"
    MozillaPath    = ".\mozilla_roots.json"
    OutputDir      = ".\output"
    BaselineObj    = $null
    WhitelistObj   = $null
    MozillaObj     = $null
    CurrentScanObj = $null
    DiffResult     = $null
    RiskResult     = $null
    Stats          = @()
    FontSize       = 13
    ZoomFactor     = 1.0
    DarkTheme      = $false
    Language       = 'zh-TW'
}

if (-not $Global:StateFlags){
    $Global:StateFlags = [ordered]@{
        MozillaCompared = $false
    }
}

$Global:PreferredFont = 'Consolas'
$Global:_UiLogBuffer  = New-Object System.Collections.Generic.List[string]

$Global:KnownPublicCANamePatterns = @(
  'microsoft','digicert','globalsign','verisign','thawte','sectigo','comodp','comodo','entrust',
  'let''s encrypt','letsencrypt','isrg','google','apple','amazon','godaddy','identrust','quovadis',
  'buypass','twca','trustwave','wotrust','wosign','gdca','harica','actalis','telia','starfield',
  'certum','swisssign','secom','hkpost','network solutions','geotrust','symantec','dtrust','lawtrust'
)

function Get-TimeTag { (Get-Date).ToString("HH:mm:ss") }

function UI-Log {
    param(
        [AllowNull()][AllowEmptyString()][string]$Message,
        [string]$Key,
        [Parameter(ValueFromRemainingArguments=$true)]
        [object[]]$Args
    )
    while ($Args -and $Args.Count -eq 1 -and $Args[0] -is [System.Array]) { $Args = @($Args[0]) }
    if ($Key) {
        $msg = L $Key
        if ($Args -and $msg -match '\{[0-9]+\}') {
            try { $msg = $msg -f $Args } catch { }
        }
        $Message = $msg
    }
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $line = "[{0}] {1}" -f (Get-TimeTag), $Message
    if ($Global:TbLog) {
        $Global:TbLog.AppendText($line + [Environment]::NewLine)
        $Global:TbLog.ScrollToEnd()
    } else {
        $Global:_UiLogBuffer.Add($line) | Out-Null
    }
    Write-Host $line
}

function Save-JsonFile {
    param([Parameter(Mandatory)]$Object,[Parameter(Mandatory)][string]$Path,[int]$Depth = 10)
    try {
        $dir=Split-Path -Parent $Path
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        ($Object | ConvertTo-Json -Depth $Depth) | Set-Content -Encoding UTF8 -Path $Path
        return $true
    } catch {
        UI-Log ("Save-JsonFile 失敗: {0}" -f $_.Exception.Message); return $false
    }
}
function Load-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { Get-Content -Raw -Encoding UTF8 -Path $Path | ConvertFrom-Json } catch {
        UI-Log ("Load-JsonFile 失敗: {0}" -f $_.Exception.Message); return $null
    }
}

# 語言資源 (含新增鍵)
$Global:LangResources = @{
  'zh-TW' = @{
    AppTitle="Root Store Viewer"; BtnLoadBaseline="載入Baseline"; BtnSaveBaseline="保存Baseline(使用目前掃描)"; BtnLoadWhitelist="載入白名單"; BtnEditWhitelist="編輯白名單"; BtnLoadMozilla="載入Mozilla"; BtnUpdateMozillaOnline="更新Mozilla(線上)"; BtnScan="掃描本機根憑證"; BtnLocalUpdate="本機更新"; BtnDiff="差異分析"; BtnRiskClassify="風險分類"; BtnMozillaCompare="Mozilla比對"; BtnHelp="說明"; BtnViewLocalUnique="本機獨有"; BtnViewMozillaUnique="Mozilla獨有"; BtnViewAdded="新增(相對Baseline)"; BtnViewRemoved="移除(相對Baseline)"; BtnViewReplaced="替換(相對Baseline)"; BtnViewExpired="過期"; BtnViewSoonExpire="即將過期"; BtnViewLongValid="長效期"; BtnZoomIn="放大"; BtnZoomOut="縮小"; BtnToggleTheme="主題切換"; BtnCreateOutput="建立Output"; LabelBaseline="基準"; LabelWhitelist="白名單"; LabelMozilla="Mozilla"; LabelOutput="輸出"; StatsTitle="統計"; LogTitle="執行日誌"; StatsColName="名稱"; StatsColCount="數量"; StatTotalCurrent="目前本機"; StatTotalBaseline="Baseline"; StatTotalMozilla="Mozilla"; StatAdded="新增"; StatRemoved="移除"; StatReplaced="替換"; StatExpired="過期"; StatSoonExpire="即將過期"; StatLongValid="長效期"; StatLocalNotInMozilla="本機≠Mozilla"; StatMozillaNotInLocal="Mozilla≠本機"; ColRiskStatus="風險狀態"; ColCategory="CA分類"; ColDaysToExpire="剩餘天數"; ColValidDays="有效期(天)"; ColSubject="主體"; ColIssuer="發行者"; ColSHA256="SHA256"; ColSerial="序號"; ColNotBefore="起始"; ColNotAfter="到期"; ColPathLength="PathLen"; ColFriendlyName="顯示名稱"; ColKeyAlgo="金鑰演算法"; ColSigAlgo="簽章演算法"; ColStoreLocation="存放位置"; ColStoreName="存放名稱"; ColEKU="EKU"; ColThumbprint="Thumbprint"; ColBase64Raw="Base64Raw"; ListWinBtnSearch="搜尋"; ListWinBtnReset="重置"; ListWinBtnSelectAll="全選"; ListWinBtnAddWhitelist="加入白名單"; ListWinBtnRemoveWhitelist="移除白名單"; LogWhitelistRemove="移除白名單 {0} 筆"; ListWinBtnExportJSON="匯出JSON"; ListWinBtnExportCSV="匯出CSV"; ListWinBtnExportDER="匯出DER"; ListWinBtnClose="關閉"; ListWinStatusTotalFormat="共 {0} 筆"; ListWinStatusSearchFormat="搜尋「{0}」 => {1}/{2}"; ListWinFilterPlaceholder="輸入關鍵字 Enter 搜尋"; RiskExpired="過期"; RiskSoon="即將過期"; RiskNormal="正常"; CatWhitelist="白名單"; CatPublic="公認"; CatUnusual="並不常見"; CatOther="其他"; HelpTitle="使用說明 / 幫助"; HelpIntro="分析 Windows 根憑證並與 Baseline / Mozilla 比較。"; HelpStartupFlow="啟動流程："; HelpStep1="1. 若無 baseline_roots.json 則掃描建立。"; HelpStep2="2. 若無 mozilla_roots.json 嘗試下載或手動匯入。"; HelpStep3="3. 載入或建立 whitelist.json。"; HelpStep4="4. 使用『本機更新』重新掃描並比較差異。"; HelpStep5="5. 檢視 / 匯出 / 新增白名單。"; HelpLangSwitch="語言下拉可即時切換。"; HelpBaselineInfo="Baseline：初次快照供 Diff。"; HelpMozillaInfo="Mozilla：Firefox 信任根清單。"; HelpWhitelistInfo="白名單：允許/忽略指紋或主體。"; HelpDiffLogic="差異：新增 / 移除 / 替換。"; HelpRisk="風險：過期 / 即將過期 / 長效期。"; HelpShortcuts="快捷鍵：F5 Diff，F6 風險，F7 Mozilla 比對，F8 語言診斷。"; HelpOpenSource="GPLv3 授權。"; AboutTitle="關於"; AboutAuthor="作者"; AboutEmail="信箱"; AboutLicense="授權"; AboutProject="專案"; AboutVersion="版本"; AboutCreated="建立時間"; AboutRepo="原始碼"; AboutClose="關閉"; MsgLangSwitched="語言已切換為"; MsgCreatingBaseline="未發現基準，建立中..."; MsgBaselineReady="Baseline 已建立"; MsgDownloadingMozilla="嘗試下載 Mozilla 清單..."; MsgMozillaDone="Mozilla 清單載入完成"; MsgMozillaFailAsk="下載失敗，是否手動選擇檔案？"; MsgMozillaImported="Mozilla 清單已匯入"; MsgMozillaImportCancel="已取消匯入 Mozilla"; MsgWhitelistLoaded="白名單載入: {0} 筆"; MsgWhitelistCreated="已建立空白白名單"; MsgAutoInitDone="自動初始化完成"; LogBaselineLoaded="Baseline 載入成功: {0}"; LogBaselineFail="Baseline 載入失敗或不存在"; LogMozillaLoaded="Mozilla 載入: {0}"; LogMozillaMissing="Mozilla 檔不存在"; LogMozillaNotLoaded="尚未載入 Mozilla"; LogNeedBaseline="尚未有 Baseline"; LogLocalUpdateStart="[本機更新] 掃描中..."; LogLocalUpdateDiff="[本機更新] Diff: +{0} -{1} ~{2}"; LogDiffResult="差異: 新增={0} 移除={1} 替換={2}"; LogRiskResult="風險: 過期={0} 即將過期={1} 長效期={2}"; LogCompareResult="Local!=Mozilla: {0} / Mozilla!=Local: {1}"; LogNoCompare="尚未執行 Mozilla 比對"; LogScanStart="開始掃描 {0}/{1}"; LogScanDone="掃描完成，共 {0} 筆，耗時 {1} ms"; LogCreateDir="已建立資料夾: {0}"; LogOutputDirCreated="已建立輸出目錄: {0}"; LogOutputDirFail="輸出目錄建立失敗: {0}"; LogWarnNoAdmin="警告：未以系統管理員執行，可能掃描不完整。"; LogIsAdmin="已以系統管理員執行。"; LogVersionStart="版本 {0} 啟動"; LogWhitelistAdd="加入白名單 {0} 筆"; LogExportJSON="輸出 JSON {0} -> {1}"; LogExportCSV="輸出 CSV {0} -> {1}"; LogExportCSVFail="CSV 失敗: {0}"; LogExportDER="DER 匯出 {0}/{1}"; LogSearch="清單搜尋 {0} => {1}"; LogZoom="Zoom={0}"; LogLangDiagItem="[LangDiag] Items={0} SelectedValue={1}"; LogLangDiagRow="[LangDiag] #{0} Code={1} Label={2}"; LogMozillaDownloadStart="開始線上更新 Mozilla..."; LogMozillaDownloadTry="嘗試來源: {0}"; LogMozillaDownloadOK="下載成功: {0} ({1} bytes)"; LogMozillaDownloadFail="下載失敗: {0}"; LogMozillaParseFromPEM="PEM 解析憑證: {0}"; LogMozillaParseFromCertdata="certdata 解析憑證: {0}"; LogMozillaNoCertFound="未解析到任何 Mozilla 憑證"; ListWinBtnUserCertMgr="使用者憑證管理"; ListWinBtnLocalCertMgr="本機憑證管理"; LogOpenUserCertMgr="開啟使用者憑證管理"; LogOpenLocalCertMgr="開啟本機憑證管理"; LogOpenCertMgrFail="憑證管理開啟失敗: {0}";MsgCertMgrNotFound= "找不到 certmgr.msc"; MsgLocalCertMgrNotFound= "找不到 certlm.msc"; MsgPkiViewNotFound= "找不到 pkiview.msc";
    # --- 新增 ---
    LogScanAuto="啟動自動掃描: {0} 筆"; LogBaselineSaved="Baseline 已保存: {0}";
    ListWinBtnEnterprisePKI="企業 PKI"; LogOpenEnterprisePKI="開啟 企業 PKI";
    FilterWhitelist="只顯示白名單"; FilterPublic="只顯示公認"; FilterUnusual="只顯示不常見"; FilterReset="恢復全部";
    # --- 1.4.1 新增 ---
    MsgAutoRiskRun="尚未計算風險，已自動執行。"; MsgRiskExpiredEmpty="沒有過期憑證"; MsgRiskSoonEmpty="沒有即將過期憑證 (<=90天)"; MsgRiskLongEmpty="沒有長效期憑證 (>15年)";
    PromptNeedCompareTitle="尚未 Mozilla 比對"; PromptNeedCompareText="尚未執行 Mozilla 比對，是否立即比對？"; PromptCompareDoneOpen="已完成 Mozilla 比對，請再次點擊欲檢視的清單。"
  }
  'zh-CN' = @{
    AppTitle="Root Store Viewer"; BtnLoadBaseline="载入Baseline"; BtnSaveBaseline="保存Baseline(使用当前扫描)"; BtnLoadWhitelist="载入白名单"; BtnEditWhitelist="编辑白名单"; BtnLoadMozilla="载入Mozilla"; BtnUpdateMozillaOnline="更新Mozilla(在线)"; BtnScan="扫描本机根证书"; BtnLocalUpdate="本机更新"; BtnDiff="差异分析"; BtnRiskClassify="风险分类"; BtnMozillaCompare="Mozilla比对"; BtnHelp="说明"; BtnViewLocalUnique="本机独有"; BtnViewMozillaUnique="Mozilla独有"; BtnViewAdded="新增(相对Baseline)"; BtnViewRemoved="移除(相对Baseline)"; BtnViewReplaced="替换(相对Baseline)"; BtnViewExpired="过期"; BtnViewSoonExpire="即将过期"; BtnViewLongValid="长期有效"; BtnZoomIn="放大"; BtnZoomOut="缩小"; BtnToggleTheme="主题切换"; BtnCreateOutput="建立输出"; LabelBaseline="基准"; LabelWhitelist="白名单"; LabelMozilla="Mozilla"; LabelOutput="输出"; StatsTitle="统计"; LogTitle="执行日志"; StatsColName="名称"; StatsColCount="数量"; StatTotalCurrent="当前本机"; StatTotalBaseline="Baseline"; StatTotalMozilla="Mozilla"; StatAdded="新增"; StatRemoved="移除"; StatReplaced="替换"; StatExpired="过期"; StatSoonExpire="即将过期"; StatLongValid="长期有效"; StatLocalNotInMozilla="本机≠Mozilla"; StatMozillaNotInLocal="Mozilla≠本机"; ColRiskStatus="风险状态"; ColCategory="CA分类"; ColDaysToExpire="剩余天数"; ColValidDays="有效期(天)"; ColSubject="主题"; ColIssuer="发行者"; ColSHA256="SHA256"; ColSerial="序列号"; ColNotBefore="起始"; ColNotAfter="到期"; ColPathLength="路径长度"; ColFriendlyName="显示名称"; ColKeyAlgo="密钥算法"; ColSigAlgo="签名算法"; ColStoreLocation="存放位置"; ColStoreName="存放名称"; ColEKU="EKU"; ColThumbprint="拇指指纹"; ColBase64Raw="Base64原始"; ListWinBtnSearch="搜索"; ListWinBtnReset="重置"; ListWinBtnSelectAll="全选"; ListWinBtnAddWhitelist="加入白名单"; ListWinBtnRemoveWhitelist="移除白名单"; LogWhitelistRemove="移除白名单 {0} 条"; ListWinBtnExportJSON="导出JSON"; ListWinBtnExportCSV="导出CSV"; ListWinBtnExportDER="导出DER"; ListWinBtnClose="关闭"; ListWinStatusTotalFormat="共 {0} 条"; ListWinStatusSearchFormat="搜索「{0}」 => {1}/{2}"; ListWinFilterPlaceholder="输入关键字 回车 搜索"; RiskExpired="过期"; RiskSoon="即将过期"; RiskNormal="正常"; CatWhitelist="白名单"; CatPublic="公认"; CatUnusual="不常见"; CatOther="其它"; HelpTitle="使用说明 / 帮助"; HelpIntro="分析 Windows 根证书并与 Baseline / Mozilla 比较。"; HelpStartupFlow="启动流程："; HelpStep1="1. 若无 baseline_roots.json 则扫描创建。"; HelpStep2="2. 若无 mozilla_roots.json 下载或手动导入。"; HelpStep3="3. 加载或创建 whitelist.json。"; HelpStep4="4. 使用『本机更新』重新扫描与比较。"; HelpStep5="5. 查看 / 导出 / 加入白名单。"; HelpLangSwitch="语言下拉即时切换。"; HelpBaselineInfo="Baseline：首次快照。"; HelpMozillaInfo="Mozilla：Firefox 信任根列表。"; HelpWhitelistInfo="白名单：允许/忽略 指纹或主题。"; HelpDiffLogic="差异：新增 / 移除 / 替换。"; HelpRisk="风险：过期 / 即将过期 / 长期有效。"; HelpShortcuts="快捷键：F5 Diff，F6 风险，F7 Mozilla，比对 F8 语言诊断。"; HelpOpenSource="GPLv3 许可。"; AboutTitle="关于"; AboutAuthor="作者"; AboutEmail="邮箱"; AboutLicense="许可"; AboutProject="项目"; AboutVersion="版本"; AboutCreated="创建时间"; AboutRepo="源码"; AboutClose="关闭"; MsgLangSwitched="语言已切换为"; MsgCreatingBaseline="未发现基准，创建中..."; MsgBaselineReady="Baseline 创建完成"; MsgDownloadingMozilla="尝试下载 Mozilla 列表..."; MsgMozillaDone="Mozilla 列表加载完成"; MsgMozillaFailAsk="下载失败，是否手动选择文件？"; MsgMozillaImported="Mozilla 列表已导入"; MsgMozillaImportCancel="已取消导入 Mozilla"; MsgWhitelistLoaded="白名单加载: {0} 条"; MsgWhitelistCreated="已创建空白白名单"; MsgAutoInitDone="自动初始化完成"; LogBaselineLoaded="Baseline 载入成功: {0}"; LogBaselineFail="Baseline 载入失败或不存在"; LogMozillaLoaded="Mozilla 载入: {0}"; LogMozillaMissing="Mozilla 文件不存在"; LogMozillaNotLoaded="尚未载入 Mozilla"; LogNeedBaseline="尚未有 Baseline"; LogLocalUpdateStart="[本机更新] 扫描中..."; LogLocalUpdateDiff="[本机更新] Diff: +{0} -{1} ~{2}"; LogDiffResult="差异: 新增={0} 移除={1} 替换={2}"; LogRiskResult="风险: 过期={0} 即将={1} 长期={2}"; LogCompareResult="Local!=Mozilla: {0} / Mozilla!=Local: {1}"; LogNoCompare="尚未执行 Mozilla 比对"; LogScanStart="开始扫描 {0}/{1}"; LogScanDone="扫描完成，共 {0} 条，耗时 {1} ms"; LogCreateDir="已建立目录: {0}"; LogOutputDirCreated="已建立输出目录: {0}"; LogOutputDirFail="输出目录建立失败: {0}"; LogWarnNoAdmin="警告：未以管理员执行，可能扫描不完整。"; LogIsAdmin="已以管理员执行。"; LogVersionStart="版本 {0} 启动"; LogWhitelistAdd="加入白名单 {0} 条"; LogExportJSON="导出 JSON {0} -> {1}"; LogExportCSV="导出 CSV {0} -> {1}"; LogExportCSVFail="CSV 失败: {0}"; LogExportDER="DER 导出 {0}/{1}"; LogSearch="列表搜索 {0} => {1}"; LogZoom="Zoom={0}"; LogLangDiagItem="[LangDiag] Items={0} SelectedValue={1}"; LogLangDiagRow="[LangDiag] #{0} Code={1} Label={2}"; LogMozillaDownloadStart="开始在线更新 Mozilla..."; LogMozillaDownloadTry="尝试来源: {0}"; LogMozillaDownloadOK="下载成功: {0} ({1} bytes)"; LogMozillaDownloadFail="下载失败: {0}"; LogMozillaParseFromPEM="PEM 解析证书: {0}"; LogMozillaParseFromCertdata="certdata 解析证书: {0}"; LogMozillaNoCertFound="未解析到任何 Mozilla 证书"; ListWinBtnUserCertMgr="用户证书管理"; ListWinBtnLocalCertMgr="本机证书管理"; LogOpenUserCertMgr="打开用户证书管理"; LogOpenLocalCertMgr="打开本机证书管理"; LogOpenCertMgrFail="证书管理打开失败: {0}"; MsgCertMgrNotFound= "未找到 certmgr.msc"; MsgLocalCertMgrNotFound= "未找到 certlm.msc"; MsgPkiViewNotFound= "未找到 pkiview.msc";
    LogScanAuto="启动自动扫描: {0} 条"; LogBaselineSaved="Baseline 已保存: {0}";
    ListWinBtnEnterprisePKI="企业 PKI"; LogOpenEnterprisePKI="打开 企业 PKI";
    FilterWhitelist="仅显示白名单"; FilterPublic="仅显示公认"; FilterUnusual="仅显示不常见"; FilterReset="恢复全部";
    MsgAutoRiskRun="尚未计算风险，已自动执行。"; MsgRiskExpiredEmpty="没有过期证书"; MsgRiskSoonEmpty="没有即将过期证书 (<=90天)"; MsgRiskLongEmpty="没有长期有效证书 (>15年)";
    PromptNeedCompareTitle="尚未 Mozilla 比对"; PromptNeedCompareText="尚未执行 Mozilla 比对，是否立即执行？"; PromptCompareDoneOpen="已完成 Mozilla 比对，请再次点击要查看的列表。"
  }
  'en' = @{
    AppTitle="Root Store Viewer"; BtnLoadBaseline="Load Baseline"; BtnSaveBaseline="Save Baseline (Use Current Scan)"; BtnLoadWhitelist="Load Whitelist"; BtnEditWhitelist="Edit Whitelist"; BtnLoadMozilla="Load Mozilla"; BtnUpdateMozillaOnline="Update Mozilla (Online)"; BtnScan="Scan Local Root Store"; BtnLocalUpdate="Local Update"; BtnDiff="Diff"; BtnRiskClassify="Risk Classify"; BtnMozillaCompare="Mozilla Diff"; BtnHelp="Help"; BtnViewLocalUnique="Local Unique"; BtnViewMozillaUnique="Mozilla Unique"; BtnViewAdded="Added (vs Baseline)"; BtnViewRemoved="Removed (vs Baseline)"; BtnViewReplaced="Replaced (vs Baseline)"; BtnViewExpired="Expired"; BtnViewSoonExpire="Soon Expire"; BtnViewLongValid="Long Valid"; BtnZoomIn="Zoom In"; BtnZoomOut="Zoom Out"; BtnToggleTheme="Toggle Theme"; BtnCreateOutput="Create Output"; LabelBaseline="Baseline"; LabelWhitelist="Whitelist"; LabelMozilla="Mozilla"; LabelOutput="Output"; StatsTitle="Statistics"; LogTitle="Log"; StatsColName="Name"; StatsColCount="Count"; StatTotalCurrent="Current Local"; StatTotalBaseline="Baseline"; StatTotalMozilla="Mozilla"; StatAdded="Added"; StatRemoved="Removed"; StatReplaced="Replaced"; StatExpired="Expired"; StatSoonExpire="Soon Expire"; StatLongValid="Long Valid"; StatLocalNotInMozilla="Local≠Mozilla"; StatMozillaNotInLocal="Mozilla≠Local"; ColRiskStatus="Risk Status"; ColCategory="Category"; ColDaysToExpire="Days Left"; ColValidDays="Valid Days"; ColSubject="Subject"; ColIssuer="Issuer"; ColSHA256="SHA256"; ColSerial="Serial"; ColNotBefore="NotBefore"; ColNotAfter="NotAfter"; ColPathLength="PathLen"; ColFriendlyName="Friendly Name"; ColKeyAlgo="Key Algo"; ColSigAlgo="Signature Algo"; ColStoreLocation="Store Location"; ColStoreName="Store Name"; ColEKU="EKU"; ColThumbprint="Thumbprint"; ColBase64Raw="Base64 Raw"; ListWinBtnSearch="Search"; ListWinBtnReset="Reset"; ListWinBtnSelectAll="Select All"; ListWinBtnAddWhitelist="Add Whitelist"; ListWinBtnRemoveWhitelist="Remove Whitelist"; LogWhitelistRemove="Removed whitelist entries {0}"; ListWinBtnExportJSON="Export JSON"; ListWinBtnExportCSV="Export CSV"; ListWinBtnExportDER="Export DER"; ListWinBtnClose="Close"; ListWinStatusTotalFormat="Total {0}"; ListWinStatusSearchFormat="Search '{0}' => {1}/{2}"; ListWinFilterPlaceholder="Type keyword Enter to search"; RiskExpired="Expired"; RiskSoon="Soon"; RiskNormal="Normal"; CatWhitelist="Whitelist"; CatPublic="Public"; CatUnusual="Unusual"; CatOther="Other"; HelpTitle="Help / Instructions"; HelpIntro="Analyze Windows root stores; compare with Baseline / Mozilla."; HelpStartupFlow="Startup flow:"; HelpStep1="1. Build baseline_roots.json if missing."; HelpStep2="2. Acquire mozilla_roots.json (download/import)."; HelpStep3="3. Load or create whitelist.json."; HelpStep4="4. Use 'Local Update' to rescan & diff."; HelpStep5="5. View / export / whitelist."; HelpLangSwitch="Switch language via combo."; HelpBaselineInfo="Baseline: initial snapshot."; HelpMozillaInfo="Mozilla: Firefox trusted root list."; HelpWhitelistInfo="Whitelist: allow/ignore entries."; HelpDiffLogic="Diff: added / removed / replaced."; HelpRisk="Risk: expired / soon / long validity."; HelpShortcuts="Shortcuts: F5 Diff, F6 Risk, F7 Mozilla Diff, F8 Lang diag."; HelpOpenSource="GPLv3 licensed."; AboutTitle="About"; AboutAuthor="Author"; AboutEmail="Email"; AboutLicense="License"; AboutProject="Project"; AboutVersion="Version"; AboutCreated="Created"; AboutRepo="Repository"; AboutClose="Close"; MsgLangSwitched="Language switched to"; MsgCreatingBaseline="Creating baseline..."; MsgBaselineReady="Baseline ready"; MsgDownloadingMozilla="Downloading Mozilla list..."; MsgMozillaDone="Mozilla list loaded"; MsgMozillaFailAsk="Download failed, import manually?"; MsgMozillaImported="Mozilla list imported"; MsgMozillaImportCancel="Mozilla import canceled"; MsgWhitelistLoaded="Whitelist loaded: {0}"; MsgWhitelistCreated="Empty whitelist created"; MsgAutoInitDone="Auto initialization complete"; LogBaselineLoaded="Baseline loaded: {0}"; LogBaselineFail="Baseline load failed or missing"; LogMozillaLoaded="Mozilla loaded: {0}"; LogMozillaMissing="Mozilla file not found"; LogMozillaNotLoaded="Mozilla not loaded"; LogNeedBaseline="Baseline missing"; LogLocalUpdateStart="[Local Update] scanning..."; LogLocalUpdateDiff="[Local Update] Diff: +{0} -{1} ~{2}"; LogDiffResult="Diff: added={0} removed={1} replaced={2}"; LogRiskResult="Risk: expired={0} soon={1} long={2}"; LogCompareResult="Local!=Mozilla: {0} / Mozilla!=Local: {1}"; LogNoCompare="Mozilla compare not run"; LogScanStart="Scan start {0}/{1}"; LogScanDone="Scan done {0} items in {1} ms"; LogCreateDir="Directory created: {0}"; LogOutputDirCreated="Output directory created: {0}"; LogOutputDirFail="Create output directory failed: {0}"; LogWarnNoAdmin="Warning: not admin, scan may be incomplete."; LogIsAdmin="Running as administrator."; LogVersionStart="Version {0} start"; LogWhitelistAdd="Whitelist added {0}"; LogExportJSON="Export JSON {0} -> {1}"; LogExportCSV="Export CSV {0} -> {1}"; LogExportCSVFail="CSV failed: {0}"; LogExportDER="Export DER {0}/{1}"; LogSearch="List search {0} => {1}"; LogZoom="Zoom={0}"; LogLangDiagItem="[LangDiag] Items={0} SelectedValue={1}"; LogLangDiagRow="[LangDiag] #{0} Code={1} Label={2}"; LogMozillaDownloadStart="Start online Mozilla update..."; LogMozillaDownloadTry="Try source: {0}"; LogMozillaDownloadOK="Download OK: {0} ({1} bytes)"; LogMozillaDownloadFail="Download failed: {0}"; LogMozillaParseFromPEM="PEM parsed certs: {0}"; LogMozillaParseFromCertdata="certdata parsed certs: {0}"; LogMozillaNoCertFound="No Mozilla certs parsed"; ListWinBtnUserCertMgr="User Cert Manager"; ListWinBtnLocalCertMgr="Local Machine Cert Manager"; LogOpenUserCertMgr="Opened user certificate manager"; LogOpenLocalCertMgr="Opened local machine certificate manager"; LogOpenCertMgrFail="Open certificate manager failed: {0}"; MsgCertMgrNotFound= "certmgr.msc not found"; MsgLocalCertMgrNotFound= "certlm.msc not found"; MsgPkiViewNotFound= "pkiview.msc not found";
    LogScanAuto="Auto scan on startup: {0} items"; LogBaselineSaved="Baseline saved: {0}";
    ListWinBtnEnterprisePKI="Enterprise PKI"; LogOpenEnterprisePKI="Opened Enterprise PKI";
    FilterWhitelist="Whitelist Only"; FilterPublic="Public CA Only"; FilterUnusual="Unusual Only"; FilterReset="Show All";
    MsgAutoRiskRun="Risk classification not yet run. Executed automatically."; MsgRiskExpiredEmpty="No expired certificates"; MsgRiskSoonEmpty="No soon-to-expire certificates (<=90 days)"; MsgRiskLongEmpty="No long-valid certificates (>15 years)";
    PromptNeedCompareTitle="Mozilla Compare Not Run"; PromptNeedCompareText="Mozilla comparison not executed yet. Run now?"; PromptCompareDoneOpen="Mozilla comparison finished. Click again to view the list."
  }
}

function L {
    param([Parameter(Mandatory)][string]$Key,[string]$Lang)
    if (-not $Lang) { $Lang=$Global:AppState.Language }
    if (-not $Global:LangResources.ContainsKey($Lang)) { $Lang='zh-TW' }
    $d=$Global:LangResources[$Lang]
    if ($d.ContainsKey($Key)) { return $d[$Key] }
    return $Key
}
#第1段落 結束

#第2段落 Baseline / Mozilla / Whitelist 與 語言設定儲存 開始
function Load-Baseline  { param([string]$Path = $Global:AppState.BaselinePath)  (Load-JsonFile -Path $Path) }
function Save-Baseline {
    param(
        [Parameter(Mandatory)]$BaselineObj,
        [string]$Path = $Global:AppState.BaselinePath
    )
    if (Save-JsonFile -Object $BaselineObj -Path $Path -Depth 12) {
        UI-Log "" "LogBaselineLoaded" @($BaselineObj.certs.Count)
        return $true
    } else {
        return $false
    }
}


function Load-Whitelist { param([string]$Path = $Global:AppState.WhitelistPath) (Load-JsonFile -Path $Path) }
function Save-Whitelist { param([Parameter(Mandatory)]$WhitelistObj,[string]$Path = $Global:AppState.WhitelistPath) if (Save-JsonFile -Object $WhitelistObj -Path $Path -Depth 6){ } }
function Load-Mozilla   { param([string]$Path = $Global:AppState.MozillaPath)   (Load-JsonFile -Path $Path) }
# --- 修正版 Save-Mozilla (回傳成功/失敗) ---
function Save-Mozilla {
    param(
        [Parameter(Mandatory)]$MozillaObj,
        [string]$Path = $Global:AppState.MozillaPath
    )
    # 可視需要調整 Depth (Mozilla certs 較多，保險給 14)
    if (Save-JsonFile -Object $MozillaObj -Path $Path -Depth 14) {
        return $true
    } else {
        return $false
    }
}

function Save-UiSettings {
    $obj=[ordered]@{
        font_size  = $Global:AppState.FontSize
        zoom       = $Global:AppState.ZoomFactor
        dark_theme = $Global:AppState.DarkTheme
        baseline   = $Global:AppState.BaselinePath
        whitelist  = $Global:AppState.WhitelistPath
        mozilla    = $Global:AppState.MozillaPath
        output_dir = $Global:AppState.OutputDir
        language   = $Global:AppState.Language
        saved_at   = (Get-Date).ToUniversalTime().ToString("o")
    }
    try {
        $json = $obj | ConvertTo-Json -Depth 4
        # UTF8 (BOM) 寫入
        [System.IO.File]::WriteAllText(
            (Join-Path (Get-Location) "gui_settings.json"),
            $json,
            [System.Text.Encoding]::UTF8
        )
    } catch {
        UI-Log ("設定儲存失敗: {0}" -f $_.Exception.Message)
    }
}
function Load-UiSettings {
    $p=".\gui_settings.json"
    if (-not (Test-Path $p)) { return }
    try {
        $cfg=Get-Content -Raw -Encoding UTF8 -Path $p|ConvertFrom-Json
        if ($cfg.font_size)   { $Global:AppState.FontSize=[int]$cfg.font_size }
        if ($cfg.zoom)        { $Global:AppState.ZoomFactor=[double]$cfg.zoom }
        if ($cfg.dark_theme -ne $null) { $Global:AppState.DarkTheme=[bool]$cfg.dark_theme }
        if ($cfg.baseline)    { $Global:AppState.BaselinePath=$cfg.baseline }
        if ($cfg.whitelist)   { $Global:AppState.WhitelistPath=$cfg.whitelist }
        if ($cfg.mozilla)     { $Global:AppState.MozillaPath=$cfg.mozilla }
        if ($cfg.output_dir)  { $Global:AppState.OutputDir=$cfg.output_dir }
        if ($cfg.language)    { $Global:AppState.Language=$cfg.language }
    } catch { UI-Log ("讀取 GUI 設定失敗: {0}" -f $_.Exception.Message) }
}

function Get-CertSha256Hex {
    param([byte[]]$RawData)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try { ($sha256.ComputeHash($RawData) | ForEach-Object { $_.ToString("x2") }) -join '' } finally { $sha256.Dispose() }
}
function Get-CertKey {
    param($CertObj)
    if ($CertObj.sha256) { return $CertObj.sha256 }
    return ("{0}|{1}" -f $CertObj.subject,$CertObj.serial)
}
function Build-Index {
    param($Obj)
    $dict=@{}
    if (-not $Obj -or -not $Obj.certs) { return $dict }
    foreach ($c in $Obj.certs) {
        $k=Get-CertKey $c
        if (-not $dict.ContainsKey($k)) { $dict[$k]=$c }
    }
    return $dict
}
#第2段落 結束
#第3段落 掃描 / Diff / 風險 / 統計 + Ensure* 補丁 開始
function Scan-LocalRootStore {
    param([string]$StoreLocation = "LocalMachine",[string]$StoreName = "Root")
    UI-Log "" "LogScanStart" @($StoreLocation,$StoreName)
    $list=@(); $sw=[System.Diagnostics.Stopwatch]::StartNew()
    try {
        $locEnum =[System.Security.Cryptography.X509Certificates.StoreLocation]::$StoreLocation
        $nameEnum=[System.Security.Cryptography.X509Certificates.StoreName]::$StoreName
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($nameEnum,$locEnum)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        foreach ($cert in $store.Certificates) {
            $raw = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
            $sha256 = Get-CertSha256Hex -RawData $raw
            $isCA=$false;$pathLen=$null;$eku=@()
            try {
                foreach ($ext in $cert.Extensions) {
                    if ($ext.Oid.Value -eq "2.5.29.19") {
                        $raw2=New-Object System.Security.Cryptography.AsnEncodedData($ext.Oid,$ext.RawData)
                        $txt=$raw2.Format($true)
                        if ($txt -match "Subject Type=.*CA") { $isCA=$true }
                        if ($txt -match "Path Length Constraint=(\d+)") { $pathLen=[int]$Matches[1] }
                    } elseif ($ext.Oid.Value -eq "2.5.29.37") {
                        $raw2=New-Object System.Security.Cryptography.AsnEncodedData($ext.Oid,$ext.RawData)
                        $eku += ($raw2.Format($true) -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                }
            } catch {}
            $isRoot = ($cert.Subject -eq $cert.Issuer)
            $item=[ordered]@{
                subject       = $cert.Subject
                issuer        = $cert.Issuer
                sha256        = $sha256
                thumbprint    = $cert.Thumbprint
                serial        = ($cert.SerialNumber -replace '\s','')
                not_before    = $cert.NotBefore.ToString("o")
                not_after     = $cert.NotAfter.ToString("o")
                friendly_name = $cert.FriendlyName
                key_algo      = $cert.PublicKey.Oid.FriendlyName
                signature_algo= $cert.SignatureAlgorithm.FriendlyName
                eku           = $eku
                is_ca         = $isCA
                is_root       = $isRoot
                path_length   = $pathLen
                store_location= $StoreLocation
                store_name    = $StoreName
                base64_raw    = [Convert]::ToBase64String($raw)
            }
            $list += [pscustomobject]$item
        }
        $store.Close()
    } catch {
        UI-Log ("掃描失敗: {0}" -f $_.Exception.Message)
    }
    $sw.Stop()
    UI-Log "" "LogScanDone" @($list.Count,$sw.ElapsedMilliseconds)
    [pscustomobject]@{
        scanned_at=(Get-Date).ToString("o")
        count     =$list.Count
        certs     =$list
        location  =$StoreLocation
        name      =$StoreName
    }
}

function Compute-RootStoreDiff {
    param([Parameter(Mandatory)]$BaselineObj,[Parameter(Mandatory)]$CurrentObj)
    if (-not $BaselineObj.certs -or -not $CurrentObj.certs) { UI-Log "Diff: 資料不完整"; return $null }
    $bIndex=Build-Index -Obj $BaselineObj
    $cIndex=Build-Index -Obj $CurrentObj
    $added=@();$removed=@();$replaced=@()
    foreach ($k in $cIndex.Keys){ if (-not $bIndex.ContainsKey($k)){ $added += $cIndex[$k] } }
    foreach ($k in $bIndex.Keys){ if (-not $cIndex.ContainsKey($k)){ $removed += $bIndex[$k] } }
    foreach ($k in $bIndex.Keys){
        if ($cIndex.ContainsKey($k)){
            $b=$bIndex[$k];$c=$cIndex[$k]
            if ($b.not_after -ne $c.not_after -or $b.issuer -ne $c.issuer){
                $replaced += [pscustomobject]@{ key=$k; baseline=$b; current=$c }
            }
        }
    }
    [pscustomobject]@{
        baseline_count=$BaselineObj.certs.Count
        current_count =$CurrentObj.certs.Count
        added_count   =$added.Count
        removed_count =$removed.Count
        replaced_count=$replaced.Count
        added=$added; removed=$removed; replaced=$replaced
        computed_at=(Get-Date).ToString("o")
    }
}

function Classify-Risks {
    param([Parameter(Mandatory)]$CurrentObj,[int]$SoonDays=90,[int]$LongYears=15)
    if (-not $CurrentObj.certs){return $null}
    $now=Get-Date
    $expired=@();$soon=@();$long=@()
    foreach ($c in $CurrentObj.certs){
        $na=Get-Date $c.not_after
        if ($na -lt $now){$expired+=$c;continue}
        if ($na -lt $now.AddDays($SoonDays)){$soon+=$c}
        $nb=Get-Date $c.not_before
        if (($na - $nb).TotalDays -gt ($LongYears*365)){$long+=$c}
    }
    [pscustomobject]@{
        expired_count=$expired.Count
        soon_expire_count=$soon.Count
        long_valid_count=$long.Count
        expired=$expired
        soon_expire=$soon
        long_valid=$long
        classified_at=(Get-Date).ToString("o")
    }
}

# --- 1.4.1 新增 Ensure 系列 ---
function Ensure-RiskResult {
    if (-not $Global:AppState.CurrentScanObj){
        $Global:AppState.CurrentScanObj = Scan-LocalRootStore
    }
    if (-not $Global:AppState.RiskResult){
        $Global:AppState.RiskResult = Classify-Risks -CurrentObj $Global:AppState.CurrentScanObj
        if ($Global:AppState.RiskResult){
            UI-Log "" "MsgAutoRiskRun"
            UI-Log "" "LogRiskResult" @(
                $Global:AppState.RiskResult.expired_count,
                $Global:AppState.RiskResult.soon_expire_count,
                $Global:AppState.RiskResult.long_valid_count
            )
            Update-Stats
        }
    }
    return $Global:AppState.RiskResult
}

function Ensure-MozillaCompared {
    if ($Global:StateFlags.MozillaCompared){ return $true }
    if (-not $Global:AppState.MozillaObj){
        UI-Log "" "LogMozillaNotLoaded"
        return $false
    }
    if (-not $Global:AppState.CurrentScanObj){
        $Global:AppState.CurrentScanObj = Scan-LocalRootStore
    }
    $ans = [System.Windows.MessageBox]::Show((L 'PromptNeedCompareText'),(L 'PromptNeedCompareTitle'),'YesNo','Question')
    if ($ans -eq 'Yes'){
        $mozIdx=Build-Index -Obj $Global:AppState.MozillaObj
        $locIdx=Build-Index -Obj $Global:AppState.CurrentScanObj
        $localNot=@();$mozNot=@()
        foreach($k in $locIdx.Keys){ if (-not $mozIdx.ContainsKey($k)){ $localNot+=$locIdx[$k]} }
        foreach($k in $mozIdx.Keys){ if (-not $locIdx.ContainsKey($k)){ $mozNot+=$mozIdx[$k]} }
        $Global:_LastLocalNotInMozilla=$localNot
        $Global:_LastMozillaNotInLocal=$mozNot
        $Global:StateFlags.MozillaCompared=$true
        UI-Log "" "LogCompareResult" @($localNot.Count,$mozNot.Count)
        Update-Stats
        UI-Log "" "PromptCompareDoneOpen"
        return $true
    } else {
        UI-Log "" "LogNoCompare"
        return $false
    }
}

function Build-BaseStats {
    @(
        @{ key="StatTotalCurrent";        getter={ if ($Global:AppState.CurrentScanObj){$Global:AppState.CurrentScanObj.certs.Count}else{0} } }
        @{ key="StatTotalBaseline";       getter={ if ($Global:AppState.BaselineObj){$Global:AppState.BaselineObj.certs.Count}else{0} } }
        @{ key="StatTotalMozilla";        getter={ if ($Global:AppState.MozillaObj){$Global:AppState.MozillaObj.certs.Count}else{0} } }
        @{ key="StatLocalNotInMozilla";   getter={ if ($Global:_LastLocalNotInMozilla){$Global:_LastLocalNotInMozilla.Count}else{0} } }
        @{ key="StatMozillaNotInLocal";   getter={ if ($Global:_LastMozillaNotInLocal){$Global:_LastMozillaNotInLocal.Count}else{0} } }
        @{ key="StatAdded";               getter={ if ($Global:AppState.DiffResult){$Global:AppState.DiffResult.added_count}else{0} } }
        @{ key="StatRemoved";             getter={ if ($Global:AppState.DiffResult){$Global:AppState.DiffResult.removed_count}else{0} } }
        @{ key="StatReplaced";            getter={ if ($Global:AppState.DiffResult){$Global:AppState.DiffResult.replaced_count}else{0} } }
        @{ key="StatExpired";             getter={ if ($Global:AppState.RiskResult){$Global:AppState.RiskResult.expired_count}else{0} } }
        @{ key="StatSoonExpire";          getter={ if ($Global:AppState.RiskResult){$Global:AppState.RiskResult.soon_expire_count}else{0} } }
        @{ key="StatLongValid";           getter={ if ($Global:AppState.RiskResult){$Global:AppState.RiskResult.long_valid_count}else{0} } }
    )
}
$Global:AppState.Stats = Build-BaseStats

function Update-Stats {
    if (-not $Global:DgStats){ return }
    $rows=@()
    foreach($s in $Global:AppState.Stats){
        $val=& $s.getter
        $rows += [pscustomobject]@{
            Name = L $s.key
            Count= $val
        }
    }
    $Global:DgStats.ItemsSource=$null
    $Global:DgStats.ItemsSource=$rows
}
#第3段落 結束
#第4段落 白名單 / 風險標註 / 顯示轉換 開始
function _NZ { param($Value) if ($null -eq $Value){return ""} return [string]$Value }

function Test-WhitelistCert {
    param($CertObject)
    if (-not $Global:AppState.WhitelistObj){return $false}
    $allow=$Global:AppState.WhitelistObj.allow
    if (-not $allow){return $false}
    $sha=(_NZ $CertObject.sha256).ToLower()
    $subj=(_NZ $CertObject.subject).ToLower()
    foreach($w in $allow){
        if ($null -eq $w){continue}
        $l=$w.ToString().ToLower()
        if ($l -eq $sha -or $l -eq $subj){return $true}
    }
    return $false
}
function Test-KnownPublicCA {
    param($CertObject)
    $fields=@(
        (_NZ $CertObject.subject).ToLower(),
        (_NZ $CertObject.issuer).ToLower()
    )
    foreach($pat in $Global:KnownPublicCANamePatterns){
        $p=$pat.ToLower()
        foreach($f in $fields){ if ($f -like "*$p*"){ return $true } }
    }
    return $false
}
function Test-UnusualLocal {
    param($CertObject)
    if (Test-WhitelistCert -CertObject $CertObject){return $false}
    if (Test-KnownPublicCA -CertObject $CertObject){return $false}
    $subj=$CertObject.subject
    if (-not $subj){return $true}
    $simple=($subj -split ',')[0].Trim()
    $short=($simple -match '^CN\s*=\s*([A-Za-z0-9\.-]{0,6})$')
    $isLocalUnique=$false
    if ($Global:_LastLocalNotInMozilla){
        $hash=$CertObject.sha256
        $isLocalUnique = $Global:_LastLocalNotInMozilla | Where-Object { $_.sha256 -eq $hash } | ForEach-Object { $true } | Select-Object -First 1
    }
    return ($short -or $isLocalUnique)
}

function Convert-ToDisplayObject {
    param($Item)
    if ($Item -is [System.Collections.IDictionary]) {
        $order=@('subject','issuer','sha256','serial','not_before','not_after','thumbprint','key_algo','signature_algo','eku','friendly_name','is_ca','is_root','path_length','store_location','store_name')
        $bag=[ordered]@{}
        foreach($k in $order){ if ($Item.Contains($k)){ $bag[$k]=$Item[$k] } }
        foreach($k in $Item.Keys){ if (-not $bag.Contains($k)){ $bag[$k]=$Item[$k]} }
        return [pscustomobject]$bag
    } elseif ($Item -is [pscustomobject]) {
        return $Item
    } else {
        $props=$Item|Get-Member -MemberType NoteProperty,Property|Select-Object -ExpandProperty Name
        $o=[ordered]@{}; foreach($p in $props){$o[$p]=$Item.$p}; return [pscustomobject]$o
    }
}

function Localize-RiskAndCategoryValue {
    param([string]$Risk,[string]$Category)
    $RMap=@{
        '過期'='RiskExpired'; 'Expired'='RiskExpired'
        '即將過期'='RiskSoon'; 'Soon Expire'='RiskSoon'; 'Soon'='RiskSoon'
        '正常'='RiskNormal'; 'Normal'='RiskNormal'
    }
    $CMap=@{
        '白名單'='CatWhitelist'; 'Whitelist'='CatWhitelist'
        '公認'='CatPublic'; 'Public'='CatPublic'
        '並不常見'='CatUnusual'; 'Unusual'='CatUnusual'
        '其他'='CatOther'; 'Other'='CatOther'
    }
    $rKey = $RMap[$Risk]; $cKey=$CMap[$Category]
    if ($rKey){ $Risk = L $rKey }
    if ($cKey){ $Category = L $cKey }
    return ,$Risk + ,$Category
}

function Add-RiskAndCategoryAnnotations {
    param([System.Collections.IEnumerable]$Items)
    $now=Get-Date
    foreach($o in $Items){
        try{
            $na= if($o.not_after){[DateTime]::Parse($o.not_after)} else {$null}
            $nb= if($o.not_before){[DateTime]::Parse($o.not_before)} else {$null}
        }catch{$na=$null;$nb=$null}
        $daysToExpire=$null;$riskRaw='正常'
        if ($na){
            $daysToExpire=[math]::Round(($na - $now).TotalDays,2)
            if ($daysToExpire -lt 0){$riskRaw='過期'}
            elseif ($daysToExpire -le 90){$riskRaw='即將過期'}
        }
        $validDays=$null
        if ($na -and $nb){ $validDays=[math]::Round(($na - $nb).TotalDays,2) }
        $isSelf = ($o.is_root -eq $true)
        $isWhitelist = Test-WhitelistCert -CertObject $o
        $isPublic    = Test-KnownPublicCA -CertObject $o
        $isUnusual   = Test-UnusualLocal -CertObject $o
        $catRaw='其他'
        if ($isWhitelist){$catRaw='白名單'}
        elseif($isPublic){$catRaw='公認'}
        elseif($isUnusual){$catRaw='並不常見'}
        $riskLocalized,$catLocalized = Localize-RiskAndCategoryValue -Risk $riskRaw -Category $catRaw
        $add=@(
            @{n='risk_status';v=$riskLocalized},
            @{n='risk_status_raw';v=$riskRaw},
            @{n='days_to_expire';v=$daysToExpire},
            @{n='valid_days';v=$validDays},
            @{n='self_signed';v=$isSelf},
            @{n='ca_category';v=$catLocalized},
            @{n='ca_category_raw';v=$catRaw},
            @{n='is_whitelisted';v=$isWhitelist},
            @{n='is_public_ca';v=$isPublic},
            @{n='is_unusual_local';v=$isUnusual}
        )
        foreach($kv in $add){
            if (-not ($o.PSObject.Properties.Name -contains $kv.n)){
                Add-Member -InputObject $o -NotePropertyName $kv.n -NotePropertyValue $kv.v
            } else {
                $o.($kv.n)=$kv.v
            }
        }
    }
    return $Items
}

function Prepare-DisplayCollection {
    param([System.Collections.IEnumerable]$Items)
    if ($null -eq $Items) { return @() }
    $converted=@()
    foreach($x in $Items){ $converted += (Convert-ToDisplayObject $x) }
    if ($converted.Count -gt 0){ Add-RiskAndCategoryAnnotations -Items $converted | Out-Null }
    return $converted
}
#第4段落 結束
#第5段落 自動資料初始化 / 語言 Combo / 語言套用 / 欄位標題 更新 開始
function Initialize-DataSources {
    if (-not (Test-Path $Global:AppState.BaselinePath)){
        UI-Log "" "MsgCreatingBaseline"
        $scan=Scan-LocalRootStore
        if ($scan){
            $Global:AppState.BaselineObj=$scan
            Save-JsonFile -Object $scan -Path $Global:AppState.BaselinePath -Depth 10 | Out-Null
            UI-Log "" "MsgBaselineReady"
        }
    } else {
        $b=Load-JsonFile -Path $Global:AppState.BaselinePath
        if ($b){ $Global:AppState.BaselineObj=$b }
    }

    if (Test-Path $Global:AppState.MozillaPath){
        $mz=Load-JsonFile -Path $Global:AppState.MozillaPath
        if ($mz){$Global:AppState.MozillaObj=$mz; UI-Log "" "MsgMozillaDone"}
    } else {
        UI-Log "" "MsgDownloadingMozilla"
        $placeholder=[ordered]@{source="placeholder";fetched_at=(Get-Date).ToString("o");certs=@()}
        Save-JsonFile -Object $placeholder -Path $Global:AppState.MozillaPath -Depth 4|Out-Null
        $Global:AppState.MozillaObj=$placeholder
        UI-Log "" "MsgMozillaDone"
    }

    if (Test-Path $Global:AppState.WhitelistPath){
        $wl=Load-JsonFile -Path $Global:AppState.WhitelistPath
        if (-not $wl){ $wl=@{allow=@()}; Save-Whitelist -WhitelistObj $wl | Out-Null; UI-Log "" "MsgWhitelistCreated" }
        else {
            if (-not $wl.allow){$wl.allow=@()}
            UI-Log "" "MsgWhitelistLoaded" @($wl.allow.Count)
        }
        $Global:AppState.WhitelistObj=$wl
    } else {
        $wl=@{ allow=@() }
        Save-Whitelist -WhitelistObj $wl | Out-Null
        $Global:AppState.WhitelistObj=$wl
        UI-Log "" "MsgWhitelistCreated"
    }
    UI-Log "" "MsgAutoInitDone"
}

function Rebuild-LanguageCombo {
    if (-not $Global:CmbLanguage){ return }
    $Global:_LangItems = @(
        [pscustomobject]@{ Code='zh-TW'; Label='繁體中文' }
        [pscustomobject]@{ Code='zh-CN'; Label='简体中文' }
        [pscustomobject]@{ Code='en';    Label='English' }
    )
    $Global:CmbLanguage.ItemsSource = $Global:_LangItems
    $Global:CmbLanguage.DisplayMemberPath='Label'
    $Global:CmbLanguage.SelectedValuePath='Code'
    if ($Global:_LangItems.Code -notcontains $Global:AppState.Language){
        $Global:AppState.Language='zh-TW'
    }
    $Global:CmbLanguage.SelectedValue=$Global:AppState.Language
}

function Get-ColumnHeaderMap {
    @{
        risk_status = L 'ColRiskStatus'
        ca_category = L 'ColCategory'
        days_to_expire = L 'ColDaysToExpire'
        valid_days = L 'ColValidDays'
        subject = L 'ColSubject'
        issuer = L 'ColIssuer'
        sha256 = L 'ColSHA256'
        serial = L 'ColSerial'
        not_before = L 'ColNotBefore'
        not_after  = L 'ColNotAfter'
        path_length= L 'ColPathLength'
        friendly_name= L 'ColFriendlyName'
        key_algo = L 'ColKeyAlgo'
        signature_algo = L 'ColSigAlgo'
        store_location = L 'ColStoreLocation'
        store_name    = L 'ColStoreName'
        eku = L 'ColEKU'
        thumbprint = L 'ColThumbprint'
        base64_raw = L 'ColBase64Raw'
    }
}
function Update-ColumnHeaderLocalization { $Global:ColumnHeaderMap = Get-ColumnHeaderMap }
function Update-AllLocalizedControls {
    Update-ColumnHeaderLocalization
    if ($Global:MainWindow) { $Global:MainWindow.Title = L 'AppTitle' }
    $map = @(
      @{Var='BtnLoadBaseline';Key='BtnLoadBaseline'}
      @{Var='BtnSaveBaseline';Key='BtnSaveBaseline'}
      @{Var='BtnLoadWhitelist';Key='BtnLoadWhitelist'}
      @{Var='BtnEditWhitelist';Key='BtnEditWhitelist'}
      @{Var='BtnLoadMozilla';Key='BtnLoadMozilla'}
      @{Var='BtnUpdateMozillaOnline';Key='BtnUpdateMozillaOnline'}
      @{Var='BtnScanLocal';Key='BtnScan'}
      @{Var='BtnLocalUpdate';Key='BtnLocalUpdate'}
      @{Var='BtnDiff';Key='BtnDiff'}
      @{Var='BtnRiskClassify';Key='BtnRiskClassify'}
      @{Var='BtnMozillaCompare';Key='BtnMozillaCompare'}
      @{Var='BtnHelp';Key='BtnHelp'}
      @{Var='BtnViewLocalUnique';Key='BtnViewLocalUnique'}
      @{Var='BtnViewMozillaUnique';Key='BtnViewMozillaUnique'}
      @{Var='BtnViewAdded';Key='BtnViewAdded'}
      @{Var='BtnViewRemoved';Key='BtnViewRemoved'}
      @{Var='BtnViewReplaced';Key='BtnViewReplaced'}
      @{Var='BtnViewExpired';Key='BtnViewExpired'}
      @{Var='BtnViewSoon';Key='BtnViewSoonExpire'}
      @{Var='BtnViewLong';Key='BtnViewLongValid'}
      @{Var='BtnZoomIn';Key='BtnZoomIn'}
      @{Var='BtnZoomOut';Key='BtnZoomOut'}
      @{Var='BtnToggleTheme';Key='BtnToggleTheme'}
      @{Var='BtnCreateOutput';Key='BtnCreateOutput'}
    )
    foreach($m in $map){
        if (Get-Variable -Name $m.Var -Scope Global -ErrorAction SilentlyContinue){
            (Get-Variable -Name $m.Var -Scope Global).Value.Content = L $m.Key
        }
    }
    if ($Global:LblBaseline){ $Global:LblBaseline.Content = L 'LabelBaseline' }
    if ($Global:LblWhitelist){ $Global:LblWhitelist.Content = L 'LabelWhitelist' }
    if ($Global:LblMozilla){ $Global:LblMozilla.Content = L 'LabelMozilla' }
    if ($Global:LblOutput){ $Global:LblOutput.Content = L 'LabelOutput' }
    if ($Global:LblStatsTitle){ $Global:LblStatsTitle.Text = L 'StatsTitle' }
    if ($Global:LblLogTitle){ $Global:LblLogTitle.Text = L 'LogTitle' }
    if ($Global:DgStats -and $Global:DgStats.Columns.Count -ge 2){
        $Global:DgStats.Columns[0].Header = L 'StatsColName'
        $Global:DgStats.Columns[1].Header = L 'StatsColCount'
    }
    Update-Stats
}
function Set-UILanguage {
    param([ValidateSet('zh-TW','zh-CN','en')][string]$Language)
    $Global:AppState.Language=$Language
    UI-Log ("{0} {1}" -f (L 'MsgLangSwitched'),$Language)
    Update-AllLocalizedControls
    if ($Global:CmbLanguage){
        $Global:CmbLanguage.SelectedValue=$Language
    }
    Save-UiSettings | Out-Null
}
function Dump-LanguageComboStatus {
    if (-not $Global:CmbLanguage){ UI-Log "語言元件不存在"; return }
    $cnt = $Global:CmbLanguage.Items.Count
    $sel = $Global:CmbLanguage.SelectedValue
    UI-Log "" "LogLangDiagItem" @($cnt,$sel)
    $i=0
    foreach($it in $Global:CmbLanguage.ItemsSource){
        $i++
        UI-Log "" "LogLangDiagRow" @($i,$it.Code,$it.Label)
    }
}
#第5段落 結束
#第6段落 GUI 建構 (含下載/比對/Row5新排列) 開始
# ================================
# 第6段落：GUI 建構
# 本版包含：
#   - 視窗 / Root Grid
#   - 工具列多行 (Row1~Row6 WrapPanel)
#   - Row5 新排列 (Diff | Unique | Risk | Zoom/Theme)
#   - 主內容區 Row=1：統計 + 日誌 並排 Grid (OPT-STATS-GRID)
#   - 狀態列 Row=2
# 注意：
#   - 舊版使用 $statsPanel 垂直堆疊的程式已移除
#   - 若你仍看到 $statsPanel 相關舊碼，請刪除，僅保留本段
# ================================

$Global:MainWindow = New-Object System.Windows.Window
$Global:MainWindow.Title = "Initializing..."
$Global:MainWindow.Width  = 1400
$Global:MainWindow.Height = 920
$Global:MainWindow.WindowStartupLocation='CenterScreen'
$Global:MainWindow.WindowState='Maximized'
if ($Global:PreferredFont){ $Global:MainWindow.FontFamily = $Global:PreferredFont }

$Global:RootScaleTransform = New-Object System.Windows.Media.ScaleTransform
$rootBorder = New-Object System.Windows.Controls.Border
$rootBorder.LayoutTransform=$Global:RootScaleTransform

$rootGrid = New-Object System.Windows.Controls.Grid
$rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null  # Row 0: 工具列 (Auto)
$rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null  # Row 1: 主內容 (伸展)
$rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null  # Row 2: 狀態列 (Auto)
$rootGrid.RowDefinitions[0].Height="Auto"
$rootGrid.RowDefinitions[1].Height="*"
$rootGrid.RowDefinitions[2].Height="Auto"
$rootBorder.Child=$rootGrid
$Global:MainWindow.Content=$rootBorder

function New-Label {
    param($Text,$Margin="4,2,4,2",$FontSize=$Global:AppState.FontSize)
    $lbl=New-Object System.Windows.Controls.Label
    $lbl.Content=$Text; $lbl.Margin=$Margin; $lbl.FontSize=$FontSize
    if ($Global:PreferredFont){ $lbl.FontFamily=$Global:PreferredFont }
    return $lbl
}
function New-TextBox {
    param($Width=250,$Text="",$Margin="4,2,4,2")
    $tb=New-Object System.Windows.Controls.TextBox
    $tb.Width=$Width; $tb.Text=$Text; $tb.Margin=$Margin
    if ($Global:PreferredFont){ $tb.FontFamily=$Global:PreferredFont }
    return $tb
}
function New-Button {
    param($Text,$Click,$Margin="4,2,4,2")
    $btn=New-Object System.Windows.Controls.Button
    $btn.Content=$Text; $btn.Margin=$Margin; $btn.Padding="8,2"
    if ($Global:PreferredFont){ $btn.FontFamily=$Global:PreferredFont }
    if ($Click){ $btn.Add_Click($Click) }
    return $btn
}

# ---- 工具列 ScrollViewer (Row=0) ----
$toolScroll = New-Object System.Windows.Controls.ScrollViewer
$toolScroll.HorizontalScrollBarVisibility="Disabled"
$toolScroll.VerticalScrollBarVisibility="Auto"
$toolScroll.MaxHeight=360
[System.Windows.Controls.Grid]::SetRow($toolScroll,0)
$rootGrid.Children.Add($toolScroll)|Out-Null

$toolOuter=New-Object System.Windows.Controls.StackPanel
$toolOuter.Orientation="Vertical"
$toolScroll.Content=$toolOuter

# WrapPanel 群組
$grpRow1=New-Object System.Windows.Controls.WrapPanel; $grpRow1.Margin="2"
$grpRow2=New-Object System.Windows.Controls.WrapPanel; $grpRow2.Margin="2"
$grpRow3=New-Object System.Windows.Controls.WrapPanel; $grpRow3.Margin="2"
$grpRow4=New-Object System.Windows.Controls.WrapPanel; $grpRow4.Margin="2"
$grpRow5=New-Object System.Windows.Controls.WrapPanel; $grpRow5.Margin="2"
$grpRow6=New-Object System.Windows.Controls.WrapPanel; $grpRow6.Margin="2"
$toolOuter.Children.Add($grpRow1)|Out-Null
$toolOuter.Children.Add($grpRow2)|Out-Null
$toolOuter.Children.Add($grpRow3)|Out-Null
$toolOuter.Children.Add($grpRow4)|Out-Null
$toolOuter.Children.Add($grpRow5)|Out-Null
$toolOuter.Children.Add($grpRow6)|Out-Null

# ---- Row1 Baseline ----
$Global:LblBaseline = New-Label (L 'LabelBaseline')
$grpRow1.Children.Add($Global:LblBaseline)|Out-Null
$Global:TbBaseline=New-TextBox -Width 300 -Text $Global:AppState.BaselinePath
$grpRow1.Children.Add($Global:TbBaseline)|Out-Null
$Global:BtnLoadBaseline = New-Button -Text (L 'BtnLoadBaseline') -Click {
    $Global:AppState.BaselinePath=$Global:TbBaseline.Text
    $obj=Load-Baseline
    if ($obj){ $Global:AppState.BaselineObj=$obj; UI-Log "" "LogBaselineLoaded" @($obj.certs.Count); Update-Stats } else { UI-Log "" "LogBaselineFail" }
}
$grpRow1.Children.Add($Global:BtnLoadBaseline)|Out-Null
$Global:BtnSaveBaseline = New-Button -Text (L 'BtnSaveBaseline') -Click {
    if (-not $Global:AppState.CurrentScanObj){
        $Global:AppState.CurrentScanObj=Scan-LocalRootStore
    }
    Save-Baseline -BaselineObj $Global:AppState.CurrentScanObj
}
$grpRow1.Children.Add($Global:BtnSaveBaseline)|Out-Null

# ---- Row2 Whitelist ----
$Global:LblWhitelist = New-Label (L 'LabelWhitelist')
$grpRow2.Children.Add($Global:LblWhitelist)|Out-Null
$Global:TbWhitelist=New-TextBox -Width 300 -Text $Global:AppState.WhitelistPath
$grpRow2.Children.Add($Global:TbWhitelist)|Out-Null
$Global:BtnLoadWhitelist = New-Button -Text (L 'BtnLoadWhitelist') -Click {
    $Global:AppState.WhitelistPath=$Global:TbWhitelist.Text
    $obj=Load-Whitelist
    if ($obj){
        if (-not $obj.allow){ $obj.allow=@() }
        $Global:AppState.WhitelistObj=$obj
        UI-Log "" "MsgWhitelistLoaded" @($obj.allow.Count)
    } else {
        $Global:AppState.WhitelistObj=@{allow=@()}
        Save-Whitelist -WhitelistObj $Global:AppState.WhitelistObj | Out-Null
        UI-Log "" "MsgWhitelistCreated"
    }
}
$grpRow2.Children.Add($Global:BtnLoadWhitelist)|Out-Null
$Global:BtnEditWhitelist = New-Button -Text (L 'BtnEditWhitelist') -Click {
    if (-not (Test-Path $Global:TbWhitelist.Text)) {
        $empty=@{allow=@()}; Save-Whitelist -WhitelistObj $empty | Out-Null
    }
    notepad $Global:TbWhitelist.Text
}
$grpRow2.Children.Add($Global:BtnEditWhitelist)|Out-Null

# ---- Row3 Mozilla ----
$Global:LblMozilla = New-Label (L 'LabelMozilla')
$grpRow3.Children.Add($Global:LblMozilla)|Out-Null
$Global:TbMozilla=New-TextBox -Width 300 -Text $Global:AppState.MozillaPath
$grpRow3.Children.Add($Global:TbMozilla)|Out-Null
$Global:BtnLoadMozilla = New-Button -Text (L 'BtnLoadMozilla') -Click {
    $Global:AppState.MozillaPath=$Global:TbMozilla.Text
    if (Test-Path $Global:AppState.MozillaPath){
        $mz=Load-Mozilla
        if ($mz){ $Global:AppState.MozillaObj=$mz; UI-Log "" "LogMozillaLoaded" @($mz.certs.Count); Update-Stats }
        else { UI-Log "" "LogMozillaMissing" }
    } else { UI-Log "" "LogMozillaMissing" }
}
$grpRow3.Children.Add($Global:BtnLoadMozilla)|Out-Null

function Get-HttpText {
    param([Parameter(Mandatory)][string]$Url,[int]$TimeoutSec=60)
    try {
        UI-Log "" "LogMozillaDownloadTry" @($Url)
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSec
        if (-not $resp -or -not $resp.Content){
            UI-Log "" "LogMozillaDownloadFail" @($Url); return $null
        }
        if ($resp.Content -match '<html' -and $Url -notlike '*.html'){
            UI-Log "內容疑似 HTML (非預期) => 視為失敗"; return $null
        }
        UI-Log "" "LogMozillaDownloadOK" @($Url,$resp.Content.Length)
        return $resp.Content
    } catch {
        UI-Log "" "LogMozillaDownloadFail" @($Url); return $null
    }
}
function Convert-X509ToObject {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert,[string]$Source="Mozilla")
    $raw=$Cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $sha=Get-CertSha256Hex -RawData $raw
    $isCA=$false;$pathLen=$null;$eku=@()
    foreach($ext in $Cert.Extensions){
        try {
            if ($ext.Oid.Value -eq "2.5.29.19"){
                $asn=New-Object System.Security.Cryptography.AsnEncodedData($ext.Oid,$ext.RawData)
                $t=$asn.Format($true)
                if ($t -match "Subject Type=.*CA"){ $isCA=$true }
                if ($t -match "Path Length Constraint=(\d+)"){ $pathLen=[int]$Matches[1] }
            } elseif ($ext.Oid.Value -eq "2.5.29.37"){
                $asn=New-Object System.Security.Cryptography.AsnEncodedData($ext.Oid,$ext.RawData)
                $eku += ($asn.Format($true) -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
        } catch {}
    }
    $isRoot = ($Cert.Subject -eq $Cert.Issuer)
    [pscustomobject][ordered]@{
        subject       = $Cert.Subject
        issuer        = $Cert.Issuer
        sha256        = $sha
        thumbprint    = $Cert.Thumbprint
        serial        = ($Cert.SerialNumber -replace '\s','')
        not_before    = $Cert.NotBefore.ToString("o")
        not_after     = $Cert.NotAfter.ToString("o")
        friendly_name = $Cert.Subject
        key_algo      = $Cert.PublicKey.Oid.FriendlyName
        signature_algo= $Cert.SignatureAlgorithm.FriendlyName
        eku           = $eku
        is_ca         = $isCA
        is_root       = $isRoot
        path_length   = $pathLen
        store_location= $Source
        store_name    = "RootStore"
        base64_raw    = [Convert]::ToBase64String($raw)
    }
}
function Parse-MozillaPEM {
    param([string]$Text)
    $list=@()
    if (-not $Text){ return $list }
    $matches = [regex]::Matches($Text,'(?s)-----BEGIN CERTIFICATE-----(.*?)-----END CERTIFICATE-----')
    foreach($m in $matches){
        $body = $m.Groups[1].Value
        $b64  = ($body -split "`r?`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[A-Za-z0-9+/=]+$' }
        $joined = ($b64 -join "")
        if ($joined.Length -lt 100){ continue }
        try {
            $bytes = [Convert]::FromBase64String($joined)
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes)
            $list += (Convert-X509ToObject -Cert $cert -Source "MozillaPEM")
        } catch {}
    }
    $list
}
function Parse-MozillaCertdata {
    param([string]$Text)
    if (-not $Text){ return @() }
    $lines = $Text -split "`r?`n"
    $list=@()
    $buffer=@()
    $inBlock=$false
    $currentLabel=$null
    foreach($ln in $lines){
        if ($ln -match '^# Certificate "(.+)"'){ $currentLabel = $Matches[1] }
        if ($ln -match 'CKA_LABEL\s+UTF8\s+"(.+)"'){ $currentLabel = $Matches[1] }
        if ($ln -match 'CKA_VALUE MULTILINE_OCTAL'){ $inBlock=$true; $buffer=@(); continue }
        if ($inBlock){
            if ($ln -eq 'END'){
                $bytesList = New-Object System.Collections.Generic.List[byte]
                foreach($ol in $buffer){
                    foreach($mm in ([regex]::Matches($ol,'\\[0-7]{3}'))){
                        $oct=$mm.Value.Substring(1)
                        try { $bytesList.Add([Convert]::ToByte($oct,8)) | Out-Null } catch {}
                    }
                }
                if ($bytesList.Count -gt 0){
                    try {
                        $rawBytes = $bytesList.ToArray()
                        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawBytes)
                        $obj  = Convert-X509ToObject -Cert $cert -Source "MozillaCertdata"
                        if ($currentLabel){ $obj.friendly_name = $currentLabel }
                        $list += $obj
                    } catch {}
                }
                $inBlock=$false
                continue
            } else { $buffer += $ln }
        }
    }
    $dedup=@{};$final=@()
    foreach($c in $list){
        if (-not $dedup.ContainsKey($c.sha256)){
            $dedup[$c.sha256]=$true
            $final+=$c
        }
    }
    $final
}
function Update-MozillaOnline {
    UI-Log "" "LogMozillaDownloadStart"
    $sources = @(
        @{url="https://curl.se/ca/cacert.pem"; type="pem"}
        @{url="https://hg.mozilla.org/mozilla-central/raw-file/tip/security/nss/lib/ckfw/builtins/certdata.txt"; type="certdata"}
        @{url="https://raw.githubusercontent.com/nss-dev/nss/main/lib/ckfw/builtins/certdata.txt"; type="certdata"}
        @{url="https://hg-edge.mozilla.org/releases/mozilla-beta/file/tip/security/nss/lib/ckfw/builtins/certdata.txt"; type="certdata"}
    )
    $best=@()
    foreach($s in $sources){
        $txt = Get-HttpText -Url $s.url
        if (-not $txt){ continue }
        $parsed=@()
        if ($s.type -eq 'pem'){
            $parsed = Parse-MozillaPEM -Text $txt
            if ($parsed.Count -gt 0){ UI-Log "" "LogMozillaParseFromPEM" @($parsed.Count) }
        } else {
            $parsed = Parse-MozillaCertdata -Text $txt
            if ($parsed.Count -gt 0){ UI-Log "" "LogMozillaParseFromCertdata" @($parsed.Count) }
        }
        if ($parsed.Count -gt $best.Count){ $best = $parsed }
        if ($best.Count -ge 130){ break }
    }
    if ($best.Count -eq 0){
        UI-Log "" "LogMozillaNoCertFound"
        return $false
    }
    $obj=[ordered]@{
        source="online"
        fetched_at=(Get-Date).ToString("o")
        count=$best.Count
        certs=$best
    }
    if (Save-Mozilla -MozillaObj $obj){
        $Global:AppState.MozillaObj=$obj
        UI-Log "" "LogMozillaLoaded" @($best.Count)
        Update-Stats
        return $true
    } else {
        UI-Log "寫入 Mozilla JSON 失敗"
        return $false
    }
}
$Global:BtnUpdateMozillaOnline = New-Button -Text (L 'BtnUpdateMozillaOnline') -Click {
    if (Update-MozillaOnline){
        # 可視需要自動觸發 Ensure-MozillaCompared
    }
}
$grpRow3.Children.Add($Global:BtnUpdateMozillaOnline)|Out-Null
$Global:BtnMozillaCompare = New-Button -Text (L 'BtnMozillaCompare') -Click {
    if (-not $Global:AppState.MozillaObj){ UI-Log "" "LogMozillaNotLoaded"; return }
    if (-not $Global:AppState.CurrentScanObj){ $Global:AppState.CurrentScanObj=Scan-LocalRootStore }
    $mozIdx=Build-Index -Obj $Global:AppState.MozillaObj
    $locIdx=Build-Index -Obj $Global:AppState.CurrentScanObj
    $localNot=@();$mozNot=@()
    foreach($k in $locIdx.Keys){ if (-not $mozIdx.ContainsKey($k)){ $localNot+=$locIdx[$k]} }
    foreach($k in $mozIdx.Keys){ if (-not $locIdx.ContainsKey($k)){ $mozNot+=$mozIdx[$k]} }
    $Global:_LastLocalNotInMozilla=$localNot
    $Global:_LastMozillaNotInLocal=$mozNot
    $Global:StateFlags.MozillaCompared = $true
    UI-Log "" "LogCompareResult" @($localNot.Count,$mozNot.Count)
    Update-Stats
}
$grpRow3.Children.Add($Global:BtnMozillaCompare)|Out-Null

# ---- Row4 操作 ----
$Global:BtnScanLocal = New-Button -Text (L 'BtnScan') -Click {
    $Global:AppState.CurrentScanObj=Scan-LocalRootStore
    Update-Stats
}
$grpRow4.Children.Add($Global:BtnScanLocal)|Out-Null
$Global:BtnLocalUpdate = New-Button -Text (L 'BtnLocalUpdate') -Click {
    if (-not $Global:AppState.BaselineObj){ UI-Log "" "LogNeedBaseline"; return }
    UI-Log "" "LogLocalUpdateStart"
    $Global:AppState.CurrentScanObj=Scan-LocalRootStore
    $diff=Compute-RootStoreDiff -BaselineObj $Global:AppState.BaselineObj -CurrentObj $Global:AppState.CurrentScanObj
    if ($diff){
        $Global:AppState.DiffResult=$diff
        UI-Log "" "LogLocalUpdateDiff" @($diff.added_count,$diff.removed_count,$diff.replaced_count)
    }
    if ($Global:AppState.MozillaObj){
        $mozIdx=Build-Index -Obj $Global:AppState.MozillaObj
        $locIdx=Build-Index -Obj $Global:AppState.CurrentScanObj
        $localNot=@()
        foreach($k in $locIdx.Keys){ if (-not $mozIdx.ContainsKey($k)){ $localNot+=$locIdx[$k]} }
        $Global:_LastLocalNotInMozilla=$localNot
    }
    $risk=Classify-Risks -CurrentObj $Global:AppState.CurrentScanObj
    if ($risk){ $Global:AppState.RiskResult=$risk }
    Update-Stats
}
$grpRow4.Children.Add($Global:BtnLocalUpdate)|Out-Null
$Global:BtnDiff = New-Button -Text (L 'BtnDiff') -Click {
    if (-not $Global:AppState.BaselineObj){ UI-Log "" "LogNeedBaseline"; return }
    if (-not $Global:AppState.CurrentScanObj){ $Global:AppState.CurrentScanObj=Scan-LocalRootStore }
    $diff=Compute-RootStoreDiff -BaselineObj $Global:AppState.BaselineObj -CurrentObj $Global:AppState.CurrentScanObj
    $Global:AppState.DiffResult=$diff
    UI-Log "" "LogDiffResult" @($diff.added_count,$diff.removed_count,$diff.replaced_count)
    Update-Stats
}
$grpRow4.Children.Add($Global:BtnDiff)|Out-Null
$Global:BtnRiskClassify = New-Button -Text (L 'BtnRiskClassify') -Click {
    if (-not $Global:AppState.CurrentScanObj){ $Global:AppState.CurrentScanObj=Scan-LocalRootStore }
    $risk=Classify-Risks -CurrentObj $Global:AppState.CurrentScanObj
    $Global:AppState.RiskResult=$risk
    UI-Log "" "LogRiskResult" @($risk.expired_count,$risk.soon_expire_count,$risk.long_valid_count)
    Update-Stats
}
$grpRow4.Children.Add($Global:BtnRiskClassify)|Out-Null
$Global:CmbLanguage = New-Object System.Windows.Controls.ComboBox
$Global:CmbLanguage.Width=140
$Global:CmbLanguage.Margin='6,2,6,2'
Rebuild-LanguageCombo
$Global:CmbLanguage.Add_SelectionChanged({
    $sel=$Global:CmbLanguage.SelectedValue
    if ($sel -and $sel -ne $Global:AppState.Language){
        Set-UILanguage -Language $sel
    }
})
$grpRow4.Children.Add($Global:CmbLanguage)|Out-Null
$Global:BtnHelp = New-Button -Text (L 'BtnHelp') -Click { Show-HelpWindow }
$grpRow4.Children.Add($Global:BtnHelp)|Out-Null

# ---- Row5 新排列 (Diff | Unique | Risk | Zoom/Theme) ----
if ($grpRow5.Children){ $grpRow5.Children.Clear() }

# Diff
$Global:BtnViewAdded = New-Button -Text (L 'BtnViewAdded') -Click {
    if (-not $Global:AppState.DiffResult){ UI-Log "" "LogDiffResult" @(0,0,0); return }
    Open-CertListWindow -Title (L 'BtnViewAdded') -Items $Global:AppState.DiffResult.added
}
$grpRow5.Children.Add($Global:BtnViewAdded)|Out-Null
$Global:BtnViewRemoved = New-Button -Text (L 'BtnViewRemoved') -Click {
    if (-not $Global:AppState.DiffResult){ UI-Log "" "LogDiffResult" @(0,0,0); return }
    Open-CertListWindow -Title (L 'BtnViewRemoved') -Items $Global:AppState.DiffResult.removed
}
$grpRow5.Children.Add($Global:BtnViewRemoved)|Out-Null
$Global:BtnViewReplaced = New-Button -Text (L 'BtnViewReplaced') -Click {
    if (-not $Global:AppState.DiffResult){ UI-Log "" "LogDiffResult" @(0,0,0); return }
    $cur=@(); foreach($r in $Global:AppState.DiffResult.replaced){ $cur += $r.current }
    Open-CertListWindow -Title (L 'BtnViewReplaced') -Items $cur
}
$grpRow5.Children.Add($Global:BtnViewReplaced)|Out-Null
$grpRow5.Children.Add((New-Label " | " -Margin "2,2,2,2"))|Out-Null

# Unique
$Global:BtnViewLocalUnique = New-Button -Text (L 'BtnViewLocalUnique') -Click {
    if (-not (Ensure-MozillaCompared)){ return }
    $items = if ($Global:_LastLocalNotInMozilla){ $Global:_LastLocalNotInMozilla } else { @() }
    Open-CertListWindow -Title (L 'BtnViewLocalUnique') -Items $items -UniqueMode
}
$grpRow5.Children.Add($Global:BtnViewLocalUnique)|Out-Null
$Global:BtnViewMozillaUnique = New-Button -Text (L 'BtnViewMozillaUnique') -Click {
    if (-not (Ensure-MozillaCompared)){ return }
    $items = if ($Global:_LastMozillaNotInLocal){ $Global:_LastMozillaNotInLocal } else { @() }
    Open-CertListWindow -Title (L 'BtnViewMozillaUnique') -Items $items -UniqueMode
}
$grpRow5.Children.Add($Global:BtnViewMozillaUnique)|Out-Null
$grpRow5.Children.Add((New-Label " | " -Margin "2,2,2,2"))|Out-Null

# Risk
$Global:BtnViewExpired = New-Button -Text (L 'BtnViewExpired') -Click {
    $risk = Ensure-RiskResult
    if (-not $risk){ return }
    if ($risk.expired_count -le 0){
        UI-Log "" "MsgRiskExpiredEmpty"
        [System.Windows.MessageBox]::Show((L 'MsgRiskExpiredEmpty'))|Out-Null
        return
    }
    Open-CertListWindow -Title (L 'BtnViewExpired') -Items $risk.expired
}
$grpRow5.Children.Add($Global:BtnViewExpired)|Out-Null
$Global:BtnViewSoon = New-Button -Text (L 'BtnViewSoonExpire') -Click {
    $risk = Ensure-RiskResult
    if (-not $risk){ return }
    if ($risk.soon_expire_count -le 0){
        UI-Log "" "MsgRiskSoonEmpty"
        [System.Windows.MessageBox]::Show((L 'MsgRiskSoonEmpty'))|Out-Null
        return
    }
    Open-CertListWindow -Title (L 'BtnViewSoonExpire') -Items $risk.soon_expire
}
$grpRow5.Children.Add($Global:BtnViewSoon)|Out-Null
$Global:BtnViewLong = New-Button -Text (L 'BtnViewLongValid') -Click {
    $risk = Ensure-RiskResult
    if (-not $risk){ return }
    if ($risk.long_valid_count -le 0){
        UI-Log "" "MsgRiskLongEmpty"
        [System.Windows.MessageBox]::Show((L 'MsgRiskLongEmpty'))|Out-Null
        return
    }
    Open-CertListWindow -Title (L 'BtnViewLongValid') -Items $risk.long_valid
}
$grpRow5.Children.Add($Global:BtnViewLong)|Out-Null
$grpRow5.Children.Add((New-Label " | " -Margin "2,2,2,2"))|Out-Null

# Zoom / Theme
$Global:BtnZoomIn = New-Button -Text (L 'BtnZoomIn') -Click { $Global:AppState.ZoomFactor+=0.1; Apply-Zoom }
$grpRow5.Children.Add($Global:BtnZoomIn)|Out-Null
$Global:BtnZoomOut = New-Button -Text (L 'BtnZoomOut') -Click { $Global:AppState.ZoomFactor=[Math]::Max(0.3,$Global:AppState.ZoomFactor-0.1); Apply-Zoom }
$grpRow5.Children.Add($Global:BtnZoomOut)|Out-Null
$Global:BtnToggleTheme = New-Button -Text (L 'BtnToggleTheme') -Click { $Global:AppState.DarkTheme = -not $Global:AppState.DarkTheme; Apply-Theme }
$grpRow5.Children.Add($Global:BtnToggleTheme)|Out-Null

# ---- Row6 Output ----
$Global:LblOutput = New-Label (L 'LabelOutput')
$grpRow6.Children.Add($Global:LblOutput)|Out-Null
$Global:TbOutput=New-TextBox -Width 300 -Text $Global:AppState.OutputDir
$grpRow6.Children.Add($Global:TbOutput)|Out-Null
$Global:BtnCreateOutput = New-Button -Text (L 'BtnCreateOutput') -Click {
    if (-not (Test-Path $Global:TbOutput.Text)){
        New-Item -ItemType Directory -Path $Global:TbOutput.Text | Out-Null
        UI-Log "" "LogCreateDir" @($Global:TbOutput.Text)
    }
}
$grpRow6.Children.Add($Global:BtnCreateOutput)|Out-Null

# =======================================================
# 主內容區：統計 + 日誌 並排 Grid  (OPT-STATS-GRID)
# =======================================================
$statsLogGrid = New-Object System.Windows.Controls.Grid
$statsLogGrid.Margin='4'
$statsLogGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))|Out-Null
$statsLogGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))|Out-Null
$statsLogGrid.ColumnDefinitions[0].Width = 330
$statsLogGrid.ColumnDefinitions[1].Width = '*'
[System.Windows.Controls.Grid]::SetRow($statsLogGrid,1)
$rootGrid.Children.Add($statsLogGrid)|Out-Null

# 左：統計區
$statsLeft = New-Object System.Windows.Controls.StackPanel
$statsLeft.Orientation='Vertical'
[System.Windows.Controls.Grid]::SetColumn($statsLeft,0)
$statsLogGrid.Children.Add($statsLeft)|Out-Null

$Global:LblStatsTitle = New-Object System.Windows.Controls.TextBlock
$Global:LblStatsTitle.Text = L 'StatsTitle'
$Global:LblStatsTitle.FontWeight='Bold'
$Global:LblStatsTitle.Margin='0,0,0,4'
$statsLeft.Children.Add($Global:LblStatsTitle)|Out-Null

$Global:DgStats=New-Object System.Windows.Controls.DataGrid
$Global:DgStats.IsReadOnly=$true
$Global:DgStats.AutoGenerateColumns=$false
$Global:DgStats.HeadersVisibility='Column'
$Global:DgStats.Margin='0,0,0,6'
$Global:DgStats.CanUserAddRows=$false
$Global:DgStats.CanUserDeleteRows=$false
$Global:DgStats.SelectionMode='Single'
$Global:DgStats.GridLinesVisibility='Horizontal'
$Global:DgStats.EnableRowVirtualization = $true
$Global:DgStats.Height=260   # 固定高度；若想自動撐滿可改為 [double]::NaN
$col1=New-Object System.Windows.Controls.DataGridTextColumn
$col1.Header=(L 'StatsColName'); $col1.Binding=New-Object System.Windows.Data.Binding "Name"; $col1.Width=180
$Global:DgStats.Columns.Add($col1)|Out-Null
$col2=New-Object System.Windows.Controls.DataGridTextColumn
$col2.Header=(L 'StatsColCount'); $col2.Binding=New-Object System.Windows.Data.Binding "Count"; $col2.Width=80
$Global:DgStats.Columns.Add($col2)|Out-Null
$statsLeft.Children.Add($Global:DgStats)|Out-Null
# 若後續要加百分比欄，可在這裡追加第三欄

# 右：日誌區 (Grid: Row0 標題, Row1 TextBox)
$logRight = New-Object System.Windows.Controls.Grid
$logRight.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null
$logRight.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null
$logRight.RowDefinitions[0].Height='Auto'
$logRight.RowDefinitions[1].Height='*'
[System.Windows.Controls.Grid]::SetColumn($logRight,1)
$statsLogGrid.Children.Add($logRight)|Out-Null

$Global:LblLogTitle = New-Object System.Windows.Controls.TextBlock
$Global:LblLogTitle.Text = L 'LogTitle'
$Global:LblLogTitle.FontWeight='Bold'
$Global:LblLogTitle.Margin='0,0,0,4'
[System.Windows.Controls.Grid]::SetRow($Global:LblLogTitle,0)
$logRight.Children.Add($Global:LblLogTitle)|Out-Null

$Global:TbLog=New-Object System.Windows.Controls.TextBox
$Global:TbLog.AcceptsReturn=$true
$Global:TbLog.IsReadOnly=$true
$Global:TbLog.VerticalScrollBarVisibility='Auto'
$Global:TbLog.HorizontalScrollBarVisibility='Auto'
$Global:TbLog.FontFamily='Consolas'
$Global:TbLog.TextWrapping='NoWrap'
$Global:TbLog.Margin='0,0,0,0'
$Global:TbLog.Padding='2'
$Global:TbLog.BorderThickness='1'
$Global:TbLog.VerticalAlignment='Stretch'
$Global:TbLog.HorizontalAlignment='Stretch'
[System.Windows.Controls.Grid]::SetRow($Global:TbLog,1)
$logRight.Children.Add($Global:TbLog)|Out-Null

# 回填 UI 初期緩衝 Log
if ($Global:_UiLogBuffer -and $Global:_UiLogBuffer.Count -gt 0){
    foreach($l in $Global:_UiLogBuffer){ $Global:TbLog.AppendText($l + [Environment]::NewLine) }
    $Global:TbLog.CaretIndex=$Global:TbLog.Text.Length
    $Global:TbLog.ScrollToEnd()
    $Global:_UiLogBuffer.Clear()
}

# (可選) 之後擴充：
#   - 插入 GridSplitter：在 $statsLogGrid 第 0 欄與第 1 欄之間
#   - Log 行數限制：在 UI-Log 函式加入行數裁切 (OPT-LOGCAP)
#   - 清除 Log 按鈕 / 匯出 Log 功能
# =======================================================

# ---- 狀態列 Row=2 ----
$statusBar=New-Object System.Windows.Controls.StackPanel
$statusBar.Orientation='Horizontal'; $statusBar.Margin='4,2,4,4'
[System.Windows.Controls.Grid]::SetRow($statusBar,2)
$rootGrid.Children.Add($statusBar)|Out-Null
$lblFooter=New-Object System.Windows.Controls.TextBlock
$lblFooter.Text="Ready."
$statusBar.Children.Add($lblFooter)|Out-Null
# OPT-AUTO-RISKINIT: 啟動後立即嘗試建立本機掃描並分類風險
	if (-not $Global:AppState.CurrentScanObj) {
		try {
			$Global:AppState.CurrentScanObj = Scan-LocalRootStore
			UI-Log "" "LogScanAuto" @($Global:AppState.CurrentScanObj.certs.Count)
		} catch {
			UI-Log "啟動自動掃描失敗: $($_.Exception.Message)"
		}
	}
	if ($Global:AppState.CurrentScanObj -and -not $Global:AppState.RiskResult) {
		try {
			$risk = Classify-Risks -CurrentObj $Global:AppState.CurrentScanObj
			if ($risk){
				$Global:AppState.RiskResult = $risk
				UI-Log "" "LogRiskResult" @($risk.expired_count,$risk.soon_expire_count,$risk.long_valid_count)
			}
		} catch {
			UI-Log "啟動自動風險分類失敗: $($_.Exception.Message)"
		}
	}
	Update-Stats

UI-Log "GUI init done"
# ================================

#第6段落 結束
#第7段落 憑證列表視窗 (排版優化版 + 修正Legend與Header Style列舉 + TextAlignment/VerticalAlignment列舉化) 開始
function Build-DataGridColumns {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.DataGrid]$DataGrid,
        [AllowNull()]$SampleObjects
    )
    $DataGrid.Columns.Clear()
    if (-not $SampleObjects){ return }
    if (-not ($SampleObjects -is [System.Collections.IEnumerable])) { $SampleObjects=@($SampleObjects) }
    if ($SampleObjects.Count -eq 0){ return }

    # 收集屬性 (最多取前 8 筆樣本)
    $propsSet = New-Object System.Collections.ArrayList
    $limit=[Math]::Min(8,$SampleObjects.Count)
    for($i=0;$i -lt $limit;$i++){
        $props = ($SampleObjects[$i] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
        foreach($p in $props){
            if (-not ($propsSet -contains $p)){ [void]$propsSet.Add($p) }
        }
    }

    $preferred = @(
        'risk_status','ca_category','days_to_expire','valid_days','subject','issuer',
        'sha256','serial','not_before','not_after','self_signed'
    )
    $ordered=@()
    foreach($p in $preferred){
        if ($propsSet -contains $p){ $ordered += $p; [void]$propsSet.Remove($p) }
    }
    $ordered += @($propsSet | Sort-Object)

    $hdrMap = Get-ColumnHeaderMap

    # Header 樣式
    $styleHeader = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
    $setterAlign  = New-Object System.Windows.Setter
    $setterAlign.Property = [System.Windows.Controls.Control]::HorizontalContentAlignmentProperty
    $setterAlign.Value    = [System.Windows.HorizontalAlignment]::Center
    $styleHeader.Setters.Add($setterAlign)
    $setterWeight = New-Object System.Windows.Setter
    $setterWeight.Property = [System.Windows.Controls.Control]::FontWeightProperty
    $setterWeight.Value    = [System.Windows.FontWeights]::Bold
    $styleHeader.Setters.Add($setterWeight)

    # Element Style（TextAlignment / VerticalAlignment 全用列舉）
    $styleCenter = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $styleCenter.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.Controls.TextBlock]::TextAlignmentProperty,
        [System.Windows.TextAlignment]::Center
    )) )
    $styleCenter.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.FrameworkElement]::VerticalAlignmentProperty,
        [System.Windows.VerticalAlignment]::Center
    )) )

    $styleRight = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $styleRight.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.Controls.TextBlock]::TextAlignmentProperty,
        [System.Windows.TextAlignment]::Right
    )) )
    $styleRight.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.FrameworkElement]::VerticalAlignmentProperty,
        [System.Windows.VerticalAlignment]::Center
    )) )

    $styleLeft = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $styleLeft.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.Controls.TextBlock]::TextAlignmentProperty,
        [System.Windows.TextAlignment]::Left
    )) )
    $styleLeft.Setters.Add( (New-Object System.Windows.Setter(
        [System.Windows.FrameworkElement]::VerticalAlignmentProperty,
        [System.Windows.VerticalAlignment]::Center
    )) )

    $centerCols = @('risk_status','ca_category','self_signed')
    $rightCols  = @('days_to_expire','valid_days','path_length')

    foreach($p in $ordered){
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Binding = New-Object System.Windows.Data.Binding $p
        if ($p -in @('days_to_expire','valid_days')){ $col.Binding.StringFormat = "{}{0:0}" }
        switch($p){
            'risk_status'    { $col.Width=90 }
            'ca_category'    { $col.Width=100 }
            'days_to_expire' { $col.Width=90 }
            'valid_days'     { $col.Width=90 }
            'subject'        { $col.Width=260 }
            'issuer'         { $col.Width=240 }
            'sha256'         { $col.Width=420 }
            default          { $col.Width="SizeToCells" }
        }
        if ($hdrMap.ContainsKey($p)){ $col.Header=$hdrMap[$p] } else { $col.Header=$p }
        $col.HeaderStyle = $styleHeader
        if ($centerCols -contains $p){ $col.ElementStyle=$styleCenter }
        elseif ($rightCols -contains $p){ $col.ElementStyle=$styleRight }
        else { $col.ElementStyle=$styleLeft }
        [void]$DataGrid.Columns.Add($col)
    }
    if ($DataGrid.Columns.Count -ge 5){ $DataGrid.FrozenColumnCount=5 }
}

if (-not $Global:_CertListWindowSize){ $Global:_CertListWindowSize=@{Width=1380;Height=760} }

# LegendItem (Brush 容錯)
function New-LegendItem {
    param(
        [string]$Text,
        $Bg,
        $Fg
    )
    if (-not ($Bg -is [System.Windows.Media.Brush])){
        try {
            if ($Bg -is [string] -and $Bg){
                $conv=New-Object System.Windows.Media.BrushConverter
                $Bg=$conv.ConvertFromString($Bg)
            } else { $Bg=[System.Windows.Media.Brushes]::White }
        } catch { $Bg=[System.Windows.Media.Brushes]::White }
    }
    if (-not ($Fg -is [System.Windows.Media.Brush])){
        try {
            if ($Fg -is [string] -and $Fg){
                $conv=New-Object System.Windows.Media.BrushConverter
                $Fg=$conv.ConvertFromString($Fg)
            } else { $Fg=[System.Windows.Media.Brushes]::Black }
        } catch { $Fg=[System.Windows.Media.Brushes]::Black }
    }
    $bd=New-Object System.Windows.Controls.Border
    $bd.Background=$Bg
    $bd.BorderBrush=[System.Windows.Media.Brushes]::Gray
    $bd.BorderThickness=1
    $bd.CornerRadius=2
    $bd.Margin='0,0,6,0'
    $lbl=New-Object System.Windows.Controls.TextBlock
    $lbl.Text=$Text
    $lbl.Margin='4,1,4,1'
    $lbl.Foreground=$Fg
    $bd.Child=$lbl
    return $bd
}

function Open-CertListWindow {
    param(
        [string]$Title="Certificates",
        [AllowNull()][System.Collections.IEnumerable]$Items,
        [switch]$UniqueMode
    )
    if ($null -eq $Items){ $Items=@() }
    $displayItems = Prepare-DisplayCollection -Items $Items

    foreach($it in $displayItems){
        if ($it.PSObject.Properties.Name -contains 'days_to_expire'){
            try { $it.days_to_expire = [int]([math]::Round([double]$it.days_to_expire,0)) } catch {}
        }
        if ($it.PSObject.Properties.Name -contains 'valid_days'){
            try { $it.valid_days = [int]([math]::Round([double]$it.valid_days,0)) } catch {}
        }
    }

    $win = New-Object System.Windows.Window
    $win.Title = $Title + " (" + $displayItems.Count + ")"
    $win.Width  = $Global:_CertListWindowSize.Width
    $win.Height = $Global:_CertListWindowSize.Height
    $win.WindowStartupLocation='CenterOwner'
    $win.ResizeMode='CanResizeWithGrip'
    if ($Global:PreferredFont){ $win.FontFamily=$Global:PreferredFont }
    $win.FontSize = $Global:AppState.FontSize
    # 新增：預設最大化
    $win.WindowState = 'Maximized'

    $rootGrid = New-Object System.Windows.Controls.Grid
    $rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null
    $rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null
    $rootGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))|Out-Null
    $rootGrid.RowDefinitions[0].Height='Auto'
    $rootGrid.RowDefinitions[1].Height='*'
    $rootGrid.RowDefinitions[2].Height='Auto'

    # ===== Top 工具列 =====
    $toolBorder = New-Object System.Windows.Controls.Border
    $toolBorder.BorderBrush=[System.Windows.Media.Brushes]::Gray
    $toolBorder.BorderThickness=1
    $toolBorder.Margin='4'
    $toolBorder.Padding='4'

    $toolGrid = New-Object System.Windows.Controls.Grid
    $toolGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))|Out-Null
    $toolGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))|Out-Null
    $toolGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))|Out-Null
    $toolGrid.ColumnDefinitions[0].Width='Auto'
    $toolGrid.ColumnDefinitions[1].Width='Auto'
    $toolGrid.ColumnDefinitions[2].Width='*'

    # --- 搜尋 + 快速篩選 ---
    $panelSearch = New-Object System.Windows.Controls.StackPanel
    $panelSearch.Orientation='Vertical'
    $panelSearch.Margin='0,0,12,0'

    $searchRow = New-Object System.Windows.Controls.StackPanel
    $searchRow.Orientation='Horizontal'
    $searchRow.Margin='0,0,0,4'
    $tbFilter = New-Object System.Windows.Controls.TextBox
    $tbFilter.Width=320; $tbFilter.Margin='0,0,4,0'
    $tbFilter.ToolTip=(L 'ListWinFilterPlaceholder')
    $searchRow.Children.Add($tbFilter)|Out-Null
    $btnSearch = New-Button -Text (L 'ListWinBtnSearch') -Click { Invoke-CertListFilter }
    $searchRow.Children.Add($btnSearch)|Out-Null
    $btnResetSearch = New-Button -Text (L 'ListWinBtnReset') -Click {
        $tbFilter.Text=""; $activeFilterMode=$null; Invoke-CertListFilter
    }
    $searchRow.Children.Add($btnResetSearch)|Out-Null
    $panelSearch.Children.Add($searchRow)|Out-Null

    $filterRow = New-Object System.Windows.Controls.StackPanel
    $filterRow.Orientation='Horizontal'
    $btnFilterWhitelist = New-Button -Text (L 'FilterWhitelist') -Click { $activeFilterMode='whitelist'; Invoke-CertListFilter }
    $btnFilterPublic    = New-Button -Text (L 'FilterPublic')    -Click { $activeFilterMode='public';    Invoke-CertListFilter }
    $btnFilterUnusual   = New-Button -Text (L 'FilterUnusual')   -Click { $activeFilterMode='unusual';   Invoke-CertListFilter }
    $btnFilterReset     = New-Button -Text (L 'FilterReset')     -Click { $activeFilterMode=$null;       Invoke-CertListFilter }
    $filterRow.Children.Add($btnFilterWhitelist)|Out-Null
    $filterRow.Children.Add($btnFilterPublic)|Out-Null
    $filterRow.Children.Add($btnFilterUnusual)|Out-Null
    $filterRow.Children.Add($btnFilterReset)|Out-Null
    $panelSearch.Children.Add($filterRow)|Out-Null
    [System.Windows.Controls.Grid]::SetColumn($panelSearch,0)
    $toolGrid.Children.Add($panelSearch)|Out-Null

    # --- 操作按鈕 ---
    $panelOps = New-Object System.Windows.Controls.StackPanel
    $panelOps.Orientation='Horizontal'
    $panelOps.Margin='0,0,12,0'

    $panelOps.Children.Add((New-Button -Text (L 'ListWinBtnAddWhitelist') -Click {
        $sel=@(); foreach($x in $dg.SelectedItems){$sel+=$x}
        if ($sel.Count -eq 0){ [System.Windows.MessageBox]::Show("No Selection")|Out-Null; return }
        if (-not $Global:AppState.WhitelistObj){ $Global:AppState.WhitelistObj=@{allow=@()} }
        if (-not $Global:AppState.WhitelistObj.allow){ $Global:AppState.WhitelistObj.allow=@() }
        $added=0
        foreach($c in $sel){
            $id= if ($c.sha256){$c.sha256}else{$c.subject}
            if (-not $id){continue}
            if (-not ($Global:AppState.WhitelistObj.allow | Where-Object { $_.ToLower() -eq $id.ToLower() })){
                $Global:AppState.WhitelistObj.allow += $id
                $added++
            }
            $c.is_whitelisted=$true
            $c.ca_category=(L 'CatWhitelist')
            $c.ca_category_raw='白名單'
        }
        Save-Whitelist -WhitelistObj $Global:AppState.WhitelistObj | Out-Null
        UI-Log "" "LogWhitelistAdd" @($added)
        Refresh-RowStyle
        [System.Windows.MessageBox]::Show((L 'ListWinBtnAddWhitelist')+" $added")|Out-Null
        Invoke-CertListFilter
    }))|Out-Null

    $panelOps.Children.Add((New-Button -Text (L 'ListWinBtnRemoveWhitelist') -Click {
        $sel=@(); foreach($x in $dg.SelectedItems){$sel+=$x}
        if ($sel.Count -eq 0){ [System.Windows.MessageBox]::Show("No Selection")|Out-Null; return }
        if (-not $Global:AppState.WhitelistObj -or -not $Global:AppState.WhitelistObj.allow){
            [System.Windows.MessageBox]::Show("Whitelist empty")|Out-Null; return
        }
        $orig=@($Global:AppState.WhitelistObj.allow)
        $new=@($orig)
        $removed=0
        foreach($c in $sel){
            $idSha=$c.sha256; $idSub=$c.subject
            if (-not $idSha -and -not $idSub){ continue }
            $before=$new.Count
            $new=@($new | Where-Object {
                $_.ToLower() -ne ($idSha.ToLower()) -and $_.ToLower() -ne ($idSub.ToLower())
            })
            if ($new.Count -lt $before){
                $removed++
                $c.is_whitelisted=$false
                if ($c.is_public_ca){
                    $c.ca_category=(L 'CatPublic'); $c.ca_category_raw='公認'
                } elseif ($c.is_unusual_local){
                    $c.ca_category=(L 'CatUnusual'); $c.ca_category_raw='並不常見'
                } else {
                    $c.ca_category=(L 'CatOther'); $c.ca_category_raw='其他'
                }
            }
        }
        if ($removed -gt 0){
            $Global:AppState.WhitelistObj.allow=$new
            Save-Whitelist -WhitelistObj $Global:AppState.WhitelistObj | Out-Null
            UI-Log "" "LogWhitelistRemove" @($removed)
            Refresh-RowStyle
            [System.Windows.MessageBox]::Show((L 'ListWinBtnRemoveWhitelist')+" $removed")|Out-Null
        } else {
            [System.Windows.MessageBox]::Show((L 'ListWinBtnRemoveWhitelist')+" 0")|Out-Null
        }
        Invoke-CertListFilter
    }))|Out-Null

    $panelOps.Children.Add((New-Button -Text (L 'ListWinBtnExportJSON') -Click {
        $sel=@(); foreach($x in $dg.SelectedItems){$sel+=$x}
        if ($sel.Count -eq 0){ [System.Windows.MessageBox]::Show("No Selection")|Out-Null; return }
        $outDir=$Global:TbOutput.Text
        if (-not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir|Out-Null }
        $path=Join-Path $outDir ("export_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".json")
        ($sel|ConvertTo-Json -Depth 14)|Set-Content -Encoding UTF8 -Path $path
        UI-Log "" "LogExportJSON" @($sel.Count,$path)
    }))|Out-Null

    $panelOps.Children.Add((New-Button -Text (L 'ListWinBtnExportCSV') -Click {
        $sel=@(); foreach($x in $dg.SelectedItems){$sel+=$x}
        if ($sel.Count -eq 0){ [System.Windows.MessageBox]::Show("No Selection")|Out-Null; return }
        $outDir=$Global:TbOutput.Text
        if (-not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir|Out-Null }
        $path=Join-Path $outDir ("export_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
        try {
            $sel | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path
            UI-Log "" "LogExportCSV" @($sel.Count,$path)
        } catch {
            UI-Log "" "LogExportCSVFail" @($_.Exception.Message)
        }
    }))|Out-Null

    $panelOps.Children.Add((New-Button -Text (L 'ListWinBtnExportDER') -Click {
        $sel=@(); foreach($x in $dg.SelectedItems){$sel+=$x}
        if ($sel.Count -eq 0){ [System.Windows.MessageBox]::Show("No Selection")|Out-Null; return }
        $outDir=Join-Path $Global:TbOutput.Text ("der_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        if (-not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir|Out-Null }
        $success=0
        foreach($c in $sel){
            $raw=$null
            if ($c.base64_raw){ try { $raw=[Convert]::FromBase64String($c.base64_raw) } catch {} }
            if ($raw){
                $fn=($c.subject -replace '[^\w\.-]','_'); if (-not $fn){$fn='cert'}
                $file=Join-Path $outDir ($fn.Substring(0,[Math]::Min(50,$fn.Length)) + ".cer")
                $i=1;$base=[IO.Path]::GetFileNameWithoutExtension($file);$dir=[IO.Path]::GetDirectoryName($file)
                while(Test-Path $file){ $file=Join-Path $dir ($base + "_$i.cer"); $i++ }
                [IO.File]::WriteAllBytes($file,$raw); $success++
            }
        }
        UI-Log "" "LogExportDER" @($success,$sel.Count)
    }))|Out-Null

    [System.Windows.Controls.Grid]::SetColumn($panelOps,1)
    $toolGrid.Children.Add($panelOps)|Out-Null

    # --- 狀態 + 圖例 ---
    $statusWrap = New-Object System.Windows.Controls.StackPanel
    $statusWrap.Orientation='Vertical'

    $txtStatus = New-Object System.Windows.Controls.TextBlock
    $txtStatus.FontWeight='Bold'
    $txtStatus.Margin='0,0,0,4'
    $txtStatus.Text=([string]::Format((L 'ListWinStatusTotalFormat'),$displayItems.Count))
    $statusWrap.Children.Add($txtStatus)|Out-Null

    # 顏色 (與 RowStyle 同步)
    $brushExpired      = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,170,170))
    $brushSoon         = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,230,170))
    $brushWhitelistBg  = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(47,127,47))
    $brushWhitelistFg  = [System.Windows.Media.Brushes]::White
    $brushPublic       = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(200,245,200))
    $brushUnusual      = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,250,205))
    $brushNormal       = [System.Windows.Media.Brushes]::White

    $legendPanel = New-Object System.Windows.Controls.StackPanel
    $legendPanel.Orientation='Horizontal'
    $legendPanel.Margin='0,0,0,0'
    $legendPanel.Children.Add( (New-LegendItem (L 'RiskExpired')  $brushExpired     ([System.Windows.Media.Brushes]::Black)) )|Out-Null
    $legendPanel.Children.Add( (New-LegendItem (L 'RiskSoon')     $brushSoon        ([System.Windows.Media.Brushes]::Black)) )|Out-Null
    $legendPanel.Children.Add( (New-LegendItem (L 'CatWhitelist') $brushWhitelistBg $brushWhitelistFg) )|Out-Null
    $legendPanel.Children.Add( (New-LegendItem (L 'CatPublic')    $brushPublic      ([System.Windows.Media.Brushes]::Black)) )|Out-Null
    $legendPanel.Children.Add( (New-LegendItem (L 'CatUnusual')   $brushUnusual     ([System.Windows.Media.Brushes]::Black)) )|Out-Null
    $statusWrap.Children.Add($legendPanel)|Out-Null

    [System.Windows.Controls.Grid]::SetColumn($statusWrap,2)
    $toolGrid.Children.Add($statusWrap)|Out-Null
    $toolBorder.Child=$toolGrid
    [System.Windows.Controls.Grid]::SetRow($toolBorder,0)
    $rootGrid.Children.Add($toolBorder)|Out-Null

    # ===== DataGrid =====
    $dg = New-Object System.Windows.Controls.DataGrid
    $dg.IsReadOnly=$true
    $dg.AutoGenerateColumns=$false
    $dg.SelectionMode='Extended'
    $dg.CanUserAddRows=$false
    $dg.CanUserDeleteRows=$false
    $dg.EnableRowVirtualization=$true
    $dg.EnableColumnVirtualization=$true
    $dg.GridLinesVisibility='Horizontal'
    $dg.RowHeight=24
    $dg.AlternatingRowBackground=[System.Windows.Media.Brushes]::WhiteSmoke
    $dg.ItemsSource=$displayItems
    Build-DataGridColumns -DataGrid $dg -SampleObjects $displayItems

    # 右鍵選單
    $ctx = New-Object System.Windows.Controls.ContextMenu
    foreach($menuDef in @(
        @{Header="複製整列(JSON)"; Action={ if ($dg.SelectedItem){ ($dg.SelectedItem|ConvertTo-Json -Depth 18)|Set-Clipboard } } },
        @{Header="複製 Subject";   Action={ if ($dg.SelectedItem){ $dg.SelectedItem.subject | Set-Clipboard } } },
        @{Header="複製 SHA256";    Action={ if ($dg.SelectedItem){ $dg.SelectedItem.sha256  | Set-Clipboard } } }
    )){
        $mi=New-Object System.Windows.Controls.MenuItem
        $mi.Header=$menuDef.Header
        $mi.Add_Click($menuDef.Action)
        $ctx.Items.Add($mi)|Out-Null
    }
    $dg.ContextMenu=$ctx

    # Row 樣式
    $applyRowStyle = {
        param($row)
        if (-not $row){ return }
        $item=$row.Item
        if (-not $item){ return }
        $riskRaw=$item.risk_status_raw
        if ($riskRaw -eq '過期' -or $riskRaw -eq 'Expired'){
            $row.Background=$brushExpired; $row.Foreground=[System.Windows.Media.Brushes]::Black; return
        }
        if ($riskRaw -eq '即將過期' -or $riskRaw -eq 'Soon' -or $riskRaw -eq 'Soon Expire'){
            $row.Background=$brushSoon; $row.Foreground=[System.Windows.Media.Brushes]::Black; return
        }
        if ($item.is_whitelisted){
            $row.Background=$brushWhitelistBg; $row.Foreground=$brushWhitelistFg; return
        }
        if ($item.is_public_ca){
            $row.Background=$brushPublic; $row.Foreground=[System.Windows.Media.Brushes]::Black; return
        }
        if ($item.is_unusual_local){
            $row.Background=$brushUnusual; $row.Foreground=[System.Windows.Media.Brushes]::Black; return
        }
        $row.Background=$brushNormal; $row.Foreground=[System.Windows.Media.Brushes]::Black
    }
    function Refresh-RowStyle {
        foreach($ri in $dg.Items){
            $rowObj=$dg.ItemContainerGenerator.ContainerFromItem($ri)
            if ($rowObj){ & $applyRowStyle $rowObj }
        }
    }
    $dg.Add_LoadingRow({ param($s,$e) & $applyRowStyle $e.Row })
    $dg.Add_SelectionChanged({ Refresh-RowStyle })
    $dg.Add_MouseDoubleClick({
        if ($dg.SelectedItem){
            $obj=$dg.SelectedItem | ConvertTo-Json -Depth 14
            [System.Windows.MessageBox]::Show($obj,"JSON",'OK','None')|Out-Null
        }
    })

    [System.Windows.Controls.Grid]::SetRow($dg,1)
    $rootGrid.Children.Add($dg)|Out-Null

    # ===== 底部列 =====
    $bottomBar = New-Object System.Windows.Controls.DockPanel
    $bottomBar.Margin='4,4,4,6'
    $statusBar = New-Object System.Windows.Controls.TextBlock
    $statusBar.Margin='4,0,0,0'
    $statusBar.VerticalAlignment='Center'
    $statusBar.Text="Total: $($displayItems.Count)"
    [System.Windows.Controls.DockPanel]::SetDock($statusBar,'Left')
    $bottomBar.Children.Add($statusBar)|Out-Null

    $btnStack = New-Object System.Windows.Controls.StackPanel
    $btnStack.Orientation='Horizontal'
    $btnStack.HorizontalAlignment='Right'

    $btnStack.Children.Add((New-Button -Text (L 'ListWinBtnUserCertMgr') -Click {
        try {
            $userMsc=Join-Path $env:SystemRoot 'System32\certmgr.msc'
            if (-not (Test-Path $userMsc)){
                $alt=Join-Path $env:SystemRoot 'SysWOW64\certmgr.msc'
                if (Test-Path $alt){ $userMsc=$alt }
            }
            if (-not (Test-Path $userMsc)){ [System.Windows.MessageBox]::Show("certmgr.msc not found")|Out-Null; return }
            Start-Process -FilePath $userMsc
            UI-Log "" "LogOpenUserCertMgr" @()
        } catch {
            UI-Log "" "LogOpenCertMgrFail" @($_.Exception.Message)
            [System.Windows.MessageBox]::Show("certmgr.msc 啟動失敗: " + $_.Exception.Message)|Out-Null
        }
    }))|Out-Null
    $btnStack.Children.Add((New-Button -Text (L 'ListWinBtnLocalCertMgr') -Click {
        try {
            $localMsc=Join-Path $env:SystemRoot 'System32\certlm.msc'
            if (-not (Test-Path $localMsc)){
                $alt=Join-Path $env:SystemRoot 'SysWOW64\certlm.msc'
                if (Test-Path $alt){ $localMsc=$alt }
            }
            if (-not (Test-Path $localMsc)){ [System.Windows.MessageBox]::Show("certlm.msc not found")|Out-Null; return }
            Start-Process -FilePath $localMsc
            UI-Log "" "LogOpenLocalCertMgr" @()
        } catch {
            UI-Log "" "LogOpenCertMgrFail" @($_.Exception.Message)
            [System.Windows.MessageBox]::Show("certlm.msc 啟動失敗: " + $_.Exception.Message)|Out-Null
        }
    }))|Out-Null
    $btnStack.Children.Add((New-Button -Text (L 'ListWinBtnEnterprisePKI') -Click {
        try {
            $pkimsc=Join-Path $env:SystemRoot 'System32\pkiview.msc'
            if (-not (Test-Path $pkimsc)){ [System.Windows.MessageBox]::Show("pkiview.msc not found")|Out-Null; return }
            Start-Process -FilePath $pkimsc
            UI-Log "" "LogOpenEnterprisePKI" @()
        } catch {
            UI-Log "" "LogOpenCertMgrFail" @($_.Exception.Message)
            [System.Windows.MessageBox]::Show("pkiview.msc 啟動失敗: " + $_.Exception.Message)|Out-Null
        }
    }))|Out-Null
    $btnStack.Children.Add((New-Button -Text (L 'ListWinBtnClose') -Click { $win.Close() }))|Out-Null

    [System.Windows.Controls.DockPanel]::SetDock($btnStack,'Right')
    $bottomBar.Children.Add($btnStack)|Out-Null
    [System.Windows.Controls.Grid]::SetRow($bottomBar,2)
    $rootGrid.Children.Add($bottomBar)|Out-Null

    # ===== 篩選邏輯 =====
    $activeFilterMode = $null   # whitelist | public | unusual | $null

    function Get-BaseFiltered {
        switch ($activeFilterMode) {
            'whitelist' { return @($displayItems | Where-Object { $_.is_whitelisted }) }
            'public'    { return @($displayItems | Where-Object { $_.is_public_ca }) }
            'unusual'   { return @($displayItems | Where-Object { $_.is_unusual_local }) }
            default     { return $displayItems }
        }
    }

    function Invoke-CertListFilter {
        $kw=$tbFilter.Text.Trim().ToLower()
        $base=Get-BaseFiltered
        if (-not $kw){
            $dg.ItemsSource=$base
            $txtStatus.Text=([string]::Format((L 'ListWinStatusTotalFormat'),$base.Count))
            $statusBar.Text="Total: {0}" -f $base.Count
            return
        }
        $filtered=@()
        foreach($o in $base){
            $concat=($o | ConvertTo-Json -Depth 3 -Compress).ToLower()
            if ($concat -like "*$kw*"){ $filtered+=$o }
        }
        $dg.ItemsSource=$filtered
        $txtStatus.Text=([string]::Format((L 'ListWinStatusSearchFormat'),$kw,$filtered.Count,$base.Count))
        $statusBar.Text="Search '{0}': {1}/{2}" -f $kw,$filtered.Count,$base.Count
        UI-Log "" "LogSearch" @($kw,$filtered.Count)
    }

    $tbFilter.Add_KeyDown({ if ($_.Key -eq 'Enter'){ Invoke-CertListFilter } })
    $tbFilter.Add_TextChanged({
        if ($tbFilter.Text.Length -ge 2 -or $tbFilter.Text.Length -eq 0){
            Invoke-CertListFilter
        }
    })

    $win.Content=$rootGrid
    $win.Owner=$Global:MainWindow
    $win.Add_SizeChanged({
        $Global:_CertListWindowSize.Width =$win.Width
        $Global:_CertListWindowSize.Height=$win.Height
    })

    if (-not $Global:_CertWindows){ $Global:_CertWindows=@() }
    $entry=[pscustomobject]@{
        Window=$win
        DataGrid=$dg
        ApplyRowStyle=$applyRowStyle
    }
    $Global:_CertWindows += $entry
    $win.Add_Closed({
        if ($Global:_CertWindows){
            $Global:_CertWindows=@($Global:_CertWindows | Where-Object { $_.Window -ne $win })
        }
    })

    $win.ShowDialog()|Out-Null
}
#第7段落 結束



#第8段落 Help / Theme / 快捷鍵 / 啟動流程 開始
function Apply-Zoom {
    if ($Global:RootScaleTransform){
        $Global:RootScaleTransform.ScaleX=$Global:AppState.ZoomFactor
        $Global:RootScaleTransform.ScaleY=$Global:AppState.ZoomFactor
        UI-Log "" "LogZoom" @([string]::Format("{0:N2}",$Global:AppState.ZoomFactor))
    }
}
function Get-VisualChildren {
    param([System.Windows.DependencyObject]$Parent)
    if (-not $Parent){return}
    $count=[System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
    for($i=0;$i -lt $count;$i++){
        $child=[System.Windows.Media.VisualTreeHelper]::GetChild($Parent,$i)
        if ($child){
            $child
            Get-VisualChildren -Parent $child
        }
    }
}
function Apply-Theme {
    if (-not $Global:MainWindow){ return }

    if (-not $Global:_ThemeColors){ $Global:_ThemeColors=@{} }
    $Global:_ThemeColors.dark = @{
        windowBg        = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(30,30,30))
        windowFg        = [System.Windows.Media.Brushes]::White
        panelBg         = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(38,38,38))
        btnBg           = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(51,51,51))
        btnBgHover      = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(68,68,68))
        btnBorder       = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(85,85,85))
        tbBg            = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(45,45,48))
        tbFg            = [System.Windows.Media.Brushes]::White
        gridBg          = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(30,30,30))
        gridAlt         = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(37,37,37))
        gridHeaderBg    = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(50,50,52))
        gridHeaderFg    = [System.Windows.Media.Brushes]::White
        labelFg         = [System.Windows.Media.Brushes]::White
        borderSplitFg   = [System.Windows.Media.Brushes]::Gainsboro
        comboBg         = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(50,50,50))
        comboFg         = [System.Windows.Media.Brushes]::White
        scrollBg        = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(32,32,32))
    }
    $Global:_ThemeColors.light = @{
        windowBg        = [System.Windows.Media.Brushes]::White
        windowFg        = [System.Windows.Media.Brushes]::Black
        panelBg         = $null
        btnBg           = $null
        btnBgHover      = $null
        btnBorder       = $null
        tbBg            = $null
        tbFg            = $null
        gridBg          = $null
        gridAlt         = $null
        gridHeaderBg    = $null
        gridHeaderFg    = $null
        labelFg         = [System.Windows.Media.Brushes]::Black
        borderSplitFg   = [System.Windows.Media.Brushes]::Black
        comboBg         = $null
        comboFg         = [System.Windows.Media.Brushes]::Black
        scrollBg        = $null
    }

    $dark  = $Global:AppState.DarkTheme
    $theme = if ($dark){ $Global:_ThemeColors.dark } else { $Global:_ThemeColors.light }

    # 視窗主色
    $Global:MainWindow.Background = $theme.windowBg
    $Global:MainWindow.Foreground = $theme.windowFg

    # 初始化 hover 事件 (一次)
    if (-not $Global:_ThemeHoverHandlersInitialized) {
        $Global:_BtnMouseEnterHandler = {
            param($s,$e)
            if ($Global:AppState.DarkTheme){
                $s.Background = $Global:_ThemeColors.dark.btnBgHover
            } else {
                $s.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
            }
        }
        $Global:_BtnMouseLeaveHandler = {
            param($s,$e)
            if ($Global:AppState.DarkTheme){
                $s.Background = $Global:_ThemeColors.dark.btnBg
            } else {
                $s.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
            }
        }
        foreach($c in (Get-VisualChildren -Parent $Global:MainWindow)){
            if ($c -and $c.GetType().Name -eq 'Button' -and -not ($c.Tag -like '*ThemeHover*')){
                $c.Add_MouseEnter($Global:_BtnMouseEnterHandler) | Out-Null
                $c.Add_MouseLeave($Global:_BtnMouseLeaveHandler) | Out-Null
                $c.Tag = if ($c.Tag){ "$($c.Tag);ThemeHover" } else { 'ThemeHover' }
            }
        }
        $Global:_ThemeHoverHandlersInitialized = $true
    }

    foreach($c in (Get-VisualChildren -Parent $Global:MainWindow)){
        if (-not $c){ continue }
        $typeName=$c.GetType().Name
        switch ($typeName) {
            'WrapPanel' {
                if ($dark){ $c.Background=$theme.panelBg } else { $c.ClearValue([System.Windows.Controls.Panel]::BackgroundProperty) }
            }
            'StackPanel' {
                # 保持原本白色即可，不強制；如需可解除註解
                # if ($dark){ $c.Background=$theme.panelBg } else { $c.ClearValue([System.Windows.Controls.Panel]::BackgroundProperty) }
            }
            'Button' {
                if ($dark){
                    $c.Background = $theme.btnBg
                    $c.Foreground = [System.Windows.Media.Brushes]::White
                    $c.BorderBrush= $theme.btnBorder
                } else {
                    $c.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::BorderBrushProperty)
                }
            }
            'TextBox' {
                if ($dark){
                    $c.Background=$theme.tbBg
                    $c.Foreground=$theme.tbFg
                    $c.BorderBrush=$theme.btnBorder
                } else {
                    $c.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::BorderBrushProperty)
                }
            }
            'ComboBox' {
                if ($dark){
                    $c.Background=$theme.comboBg
                    $c.Foreground=$theme.comboFg
                } else {
                    $c.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
                }
            }
            'Label' {
                $c.Foreground = $theme.labelFg
                if ($dark){
                    # 若想讓左側標籤更凸顯，可加：$c.FontWeight='SemiBold'
                } else {
                    $c.ClearValue([System.Windows.Controls.Control]::FontWeightProperty)
                }
            }
            'TextBlock' {
                # Stats / Log 標題
                $c.Foreground = $theme.labelFg
            }
            'DataGrid' {
                if ($dark){
                    $c.Background=$theme.gridBg
                    $c.Foreground=[System.Windows.Media.Brushes]::White
                    $c.RowBackground=$theme.gridBg
                    $c.AlternatingRowBackground=$theme.gridAlt
                } else {
                    $c.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
                    $c.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
                    $c.ClearValue([System.Windows.Controls.DataGrid]::RowBackgroundProperty)
                    $c.ClearValue([System.Windows.Controls.DataGrid]::AlternatingRowBackgroundProperty)
                }
            }
        }
    }

    # DataGrid 標頭樣式 (僅統計那個)
    if ($Global:DgStats){
        if ($dark){
            $style = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Primitives.DataGridColumnHeader]::BackgroundProperty,$theme.gridHeaderBg))
            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Primitives.DataGridColumnHeader]::ForegroundProperty,$theme.gridHeaderFg))
            $Global:DgStats.ColumnHeaderStyle=$style
        } else {
            $Global:DgStats.ClearValue([System.Windows.Controls.DataGrid]::ColumnHeaderStyleProperty)
        }
    }
    # --- 主題切換後重新套用已開啟憑證視窗 Row 樣式 ---
    if ($Global:_CertWindows){
        foreach($cw in $Global:_CertWindows){
            try {
                $dg2=$cw.DataGrid
                $ap =$cw.ApplyRowStyle
                if ($dg2 -and $ap){
                    foreach($item in $dg2.Items){
                        $rowObj=$dg2.ItemContainerGenerator.ContainerFromItem($item)
                        if ($rowObj){ & $ap $rowObj }
                    }
                }
            } catch {}
        }
    }
	
}

$Global:AppMeta = @{
    Project = "Root Store Analyzer"
    Version = $Global:AppState.Version
    Created = (Get-Date -Format "yyyy-MM-dd")
    Author  = "startgo"
    Email   = "startgo@yia.app"
    License = "GPLv3"
    Repo    = "https://github.com/ystartgo"
}

function Show-HelpWindow {
    $w=New-Object System.Windows.Window
    $w.Title=L 'HelpTitle'
    $w.Width=780;$w.Height=640
    $w.WindowStartupLocation='CenterOwner'
    if ($Global:PreferredFont){$w.FontFamily=$Global:PreferredFont}
    $sp=New-Object System.Windows.Controls.StackPanel
    $sp.Margin='10'
    $tb=New-Object System.Windows.Controls.TextBlock
    $tb.TextWrapping='Wrap'
    $tb.FontFamily='Consolas';$tb.FontSize=13
    $lines=@(
        (L 'HelpIntro'),"",
        (L 'HelpStartupFlow'),
        (L 'HelpStep1'),
        (L 'HelpStep2'),
        (L 'HelpStep3'),
        (L 'HelpStep4'),
        (L 'HelpStep5'),"",
        (L 'HelpBaselineInfo'),
        (L 'HelpMozillaInfo'),
        (L 'HelpWhitelistInfo'),
        (L 'HelpDiffLogic'),
        (L 'HelpRisk'),
        (L 'HelpLangSwitch'),
        (L 'HelpShortcuts'),
        (L 'HelpOpenSource')
    )
    $tb.Text=($lines -join "`r`n")
    $sp.Children.Add($tb)|Out-Null
    $btnAbout=New-Button -Text (L 'AboutTitle') -Click { Show-AboutWindow }
    $btnAbout.Margin='0,12,0,0'
    $sp.Children.Add($btnAbout)|Out-Null
    $btnClose=New-Button -Text (L 'AboutClose') -Click { $w.Close() }
    $btnClose.Margin='6'
    $sp.Children.Add($btnClose)|Out-Null
    $w.Content=$sp; $w.Owner=$Global:MainWindow
    $w.ShowDialog()|Out-Null
}

function Show-AboutWindow {
    $w=New-Object System.Windows.Window
    $w.Title=L 'AboutTitle'
    $w.Width=520;$w.Height=430
    $w.WindowStartupLocation='CenterOwner'
    if ($Global:PreferredFont){$w.FontFamily=$Global:PreferredFont}
    $sp=New-Object System.Windows.Controls.StackPanel
    $sp.Margin='10'
    $info=@()
    $info += ("{0}: {1}" -f (L 'AboutProject'),$Global:AppMeta.Project)
    $info += ("{0}: {1}" -f (L 'AboutVersion'),$Global:AppMeta.Version)
    $info += ("{0}: {1}" -f (L 'AboutCreated'),$Global:AppMeta.Created)
    $info += ("{0}: {1}" -f (L 'AboutAuthor'),$Global:AppMeta.Author)
    $info += ("{0}: {1}" -f (L 'AboutEmail'),$Global:AppMeta.Email)
    $info += ("{0}: {1}" -f (L 'AboutLicense'),$Global:AppMeta.License)
    $info += ("{0}: {1}" -f (L 'AboutRepo'),$Global:AppMeta.Repo)
    $tb=New-Object System.Windows.Controls.TextBlock
    $tb.Text=($info -join "`r`n");$tb.FontFamily='Consolas';$tb.TextWrapping='Wrap'
    $sp.Children.Add($tb)|Out-Null
    $license=New-Object System.Windows.Controls.TextBox
    $license.IsReadOnly=$true
    $license.VerticalScrollBarVisibility='Auto'
    $license.Height=170
    $license.FontFamily='Consolas'
$license.Text=@"
GNU GENERAL PUBLIC LICENSE Version 3 (Excerpt)
This program is free software: you can redistribute it and/or modify it
under the terms of the GNU GPL v3.
Full text: https://www.gnu.org/licenses/gpl-3.0.html
"@
    $sp.Children.Add($license)|Out-Null
    $btn=New-Button -Text (L 'AboutClose') -Click { $w.Close() }
    $btn.Margin='8'
    $sp.Children.Add($btn)|Out-Null
    $w.Content=$sp; $w.Owner=$Global:MainWindow
    $w.ShowDialog()|Out-Null
}

function Test-IsAdmin {
    try {
        $id=[Security.Principal.WindowsIdentity]::GetCurrent()
        $p=New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

$Global:MainWindow.Add_KeyDown({
    if ($_.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control){
        switch ($_.Key){
            'L' {
                $obj=Load-Baseline
                if ($obj){$Global:AppState.BaselineObj=$obj;UI-Log "" "LogBaselineLoaded" @($obj.certs.Count);Update-Stats}
            }
        }
    } else {
        switch ($_.Key){
            'F5' {
                if (-not $Global:AppState.BaselineObj){UI-Log "" "LogNeedBaseline";break}
                if (-not $Global:AppState.CurrentScanObj){$Global:AppState.CurrentScanObj=Scan-LocalRootStore}
                $diff=Compute-RootStoreDiff -BaselineObj $Global:AppState.BaselineObj -CurrentObj $Global:AppState.CurrentScanObj
                $Global:AppState.DiffResult=$diff
                UI-Log "" "LogDiffResult" @($diff.added_count,$diff.removed_count,$diff.replaced_count)
                Update-Stats
            }
            'F6' {
                if (-not $Global:AppState.CurrentScanObj){$Global:AppState.CurrentScanObj=Scan-LocalRootStore}
                $risk=Classify-Risks -CurrentObj $Global:AppState.CurrentScanObj
                $Global:AppState.RiskResult=$risk
                UI-Log "" "LogRiskResult" @($risk.expired_count,$risk.soon_expire_count,$risk.long_valid_count)
                Update-Stats
            }
            'F7' {
                if (-not $Global:AppState.MozillaObj){UI-Log "" "LogMozillaNotLoaded";break}
                if (-not $Global:AppState.CurrentScanObj){$Global:AppState.CurrentScanObj=Scan-LocalRootStore}
                $mozIdx=Build-Index -Obj $Global:AppState.MozillaObj
                $locIdx=Build-Index -Obj $Global:AppState.CurrentScanObj
                $localNot=@();$mozNot=@()
                foreach($k in $locIdx.Keys){ if (-not $mozIdx.ContainsKey($k)){ $localNot+=$locIdx[$k]} }
                foreach($k in $mozIdx.Keys){ if (-not $locIdx.ContainsKey($k)){ $mozNot+=$mozIdx[$k]} }
                $Global:_LastLocalNotInMozilla=$localNot
                $Global:_LastMozillaNotInLocal=$mozNot
                $Global:StateFlags.MozillaCompared = $true
                UI-Log "" "LogCompareResult" @($localNot.Count,$mozNot.Count)
                Update-Stats
            }
            'F8' { Dump-LanguageComboStatus }
        }
    }
})

$Global:MainWindow.Add_Closing({ Save-UiSettings })

Load-UiSettings
if ($Global:AppState.Language -notin @('zh-TW','zh-CN','en')){ $Global:AppState.Language='zh-TW' }
Rebuild-LanguageCombo
Set-UILanguage -Language $Global:AppState.Language
Apply-Zoom
Apply-Theme

if ($Global:AppState.OutputDir -and -not (Test-Path $Global:AppState.OutputDir)){
    try {
        New-Item -ItemType Directory -Path $Global:AppState.OutputDir | Out-Null
        UI-Log "" "LogOutputDirCreated" @($Global:AppState.OutputDir)
    } catch {
        UI-Log "" "LogOutputDirFail" @($_.Exception.Message)
    }
}

if (-not (Test-IsAdmin)){ UI-Log "" "LogWarnNoAdmin" } else { UI-Log "" "LogIsAdmin" }

Initialize-DataSources
Update-Stats
UI-Log "" "LogVersionStart" @($Global:AppState.Version)

[void]$Global:MainWindow.ShowDialog()
UI-Log "Window closed"
#第8段落 結束
#第9段落 Wrapper / 相容函式 與 結尾 開始
function Invoke-ScanLocal {
    UI-Log "[Wrapper] Scan Local"
    $Global:AppState.CurrentScanObj=Scan-LocalRootStore
}
function Invoke-Diff {
    if (-not $Global:AppState.BaselineObj){UI-Log "" "LogNeedBaseline";return}
    if (-not $Global:AppState.CurrentScanObj){Invoke-ScanLocal}
    $diff=Compute-RootStoreDiff -BaselineObj $Global:AppState.BaselineObj -CurrentObj $Global:AppState.CurrentScanObj
    if ($diff){$Global:AppState.DiffResult=$diff;UI-Log "" "LogDiffResult" @($diff.added_count,$diff.removed_count,$diff.replaced_count);Update-Stats}
}
function Invoke-LoadBaseline {
    $obj=Load-Baseline
    if ($obj){$Global:AppState.BaselineObj=$obj;UI-Log "" "LogBaselineLoaded" @($obj.certs.Count);Update-Stats}
}
function Invoke-SaveBaseline {
    if (-not $Global:AppState.CurrentScanObj){Invoke-ScanLocal}
    Save-Baseline -BaselineObj $Global:AppState.CurrentScanObj | Out-Null
}
function Invoke-LoadMozillaManual {
    if (Test-Path $Global:AppState.MozillaPath){
        $moz=Load-Mozilla
        if ($moz){$Global:AppState.MozillaObj=$moz;UI-Log "" "LogMozillaLoaded" @($moz.certs.Count);Update-Stats}
    } else {
        UI-Log "" "LogMozillaMissing"
    }
}
function Invoke-LoadWhitelist {
    if (Test-Path $Global:AppState.WhitelistPath){
        $wl=Load-Whitelist
        if ($wl){ if (-not $wl.allow){$wl.allow=@()} $Global:AppState.WhitelistObj=$wl; UI-Log "" "MsgWhitelistLoaded" @($wl.allow.Count) }
    } else {
        UI-Log "" "MsgWhitelistCreated"
    }
}
# END OF FILE (Version 1.4.4 Patched)
#第9段落 Wrapper / 相容函式 與 結尾 結束