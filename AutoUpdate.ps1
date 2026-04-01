# [系統優化] 強制開啟 TLS 1.2 協議與隱藏下載進度
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'

# ======================================================
# Acer Marketing Auto-Update Script v1.2.1
# GitHub ID: acer-wallpaper
# ======================================================

# 1. 基礎路徑設定
$baseUrl = "https://raw.githubusercontent.com/acer-wallpaper/Acer-Wallpaper/main/"
$scriptUrl = $baseUrl + "AutoUpdate.ps1"
$localFolder = "C:\Acer_Marketing"
$localScript = "$localFolder\AutoUpdate.ps1"
$logFile = "$localFolder\sync_log.txt"

# 2. 【管理員配置】你的 Google Form 設定
$formUrl = "https://docs.google.com/forms/d/e/1FAIpQLScItcWdiC6LmnyrjpagV0_4rLV_rhL2kDHhfpR1_t2OUb1qqA/formResponse"
$entry_PCName = "entry.1791815151"
$entry_Model  = "entry.307299855"
$entry_WiFi   = "entry.286077098"
$entry_Time   = "entry.104401420"

# --- 輔助功能：寫入日誌 ---
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (!(Test-Path $localFolder)) { New-Item -ItemType Directory -Path $localFolder -Force }
    "[$timestamp] $Message" | Add-Content -Path $logFile
    if ((Get-Content $logFile).Count -gt 1000) { (Get-Content $logFile | Select-Object -Last 100) | Out-File $logFile -Encoding utf8 }
}

# --- 輔助功能：獲取 Wi-Fi SSID ---
function Get-WiFiSSID {
    try {
        $ssid = (netsh wlan show interfaces | Select-String "^\s+SSID\s+:\s+(.+)$").Matches.Groups[1].Value.Trim()
        return if ($ssid) { $ssid } else { "Ethernet" }
    } catch { return "Unknown" }
}

# --- 輔助功能：回報機台資訊 ---
function Submit-Reporting {
    param($ModelName)
    try {
        $postBody = @{
            $entry_PCName = $env:COMPUTERNAME
            $entry_Model  = $ModelName
            $entry_WiFi   = Get-WiFiSSID
            $entry_Time   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        Invoke-WebRequest -Uri $formUrl -Method Post -Body $postBody -ErrorAction SilentlyContinue -TimeoutSec 5
        Write-Log "Cloud Report: Success."
    } catch { Write-Log "Cloud Report: Failed (but continuing)." }
}

Write-Log "--- Session Started v1.2.1 ---"

# ------------------------------------------------------
# STEP A: 自我更新 (如果沒網則跳過)
# ------------------------------------------------------
try {
    # 強制確保網址中間有斜槓
    $cleanBase = $baseUrl.TrimEnd('/')
    $fullScriptUrl = "$cleanBase/AutoUpdate.ps1"
    
    $webClient = New-Object System.Net.WebClient
    $webClient.Encoding = [System.Text.Encoding]::UTF8
    $remoteContent = $webClient.DownloadString($fullScriptUrl)

    if (Test-Path $localScript) { $localContent = Get-Content $localScript -Raw }

    if ($remoteContent -and ($remoteContent -ne $localContent)) {
        Write-Log "New script detected. Updating..."
        $remoteContent | Out-File -FilePath $localScript -Encoding utf8 -Force
        # 重新啟動新腳本並結束舊腳本
        Start-Process "powershell.exe" -ArgumentList "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$localScript`"" -WindowStyle Hidden
        exit
    }
} catch { 
    # 這裡多加一行，讓我們在 LOG 看到具體錯誤
    Write-Log "Self-update skipped. Reason: $($_.Exception.Message)" 
}# ------------------------------------------------------
# STEP B: 下載與套用桌布
# ------------------------------------------------------
try {
    # 1. 取得型號並清理
    $rawModel = (Get-CimInstance Win32_ComputerSystem).Model
    $model = $rawModel.Split(' ')[-1].Trim()
    
    # 2. 回報至 Google Form (這部分你的 ID 已經通了，維持原樣)
    Submit-Reporting -ModelName $model

    $localCache = "$localFolder\latest_backup.jpg"
    $cleanBase = $baseUrl.TrimEnd('/')
    $success = $false

    # 3. 定義多種可能的檔名組合進行嘗試 (解決 404 問題)
    # 組合：原名.jpg, 原名.png, 原名.JPG, 原名.PNG
    $searchList = @("$model.jpg", "$model.png", "$model.JPG", "$model.PNG")

    foreach ($fileName in $searchList) {
        $targetUrl = "$cleanBase/$fileName"
        try {
            Write-Log "Trying to download: $targetUrl"
            Invoke-WebRequest -Uri $targetUrl -OutFile $localCache -ErrorAction Stop -TimeoutSec 10
            $success = $true
            Write-Log "SUCCESS: Found and downloaded $fileName"
            break # 找到就跳出迴圈
        } catch {
            continue # 失敗就試下一個
        }
    }

    # 4. 如果下載成功，套用桌布
    if ($success) {
        $code = @"
        using System.Runtime.InteropServices;
        public class Wallpaper {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }
"@
        if (-not ([System.Type]::GetType("Wallpaper"))) { Add-Type -TypeDefinition $code }
        [Wallpaper]::SystemParametersInfo(0x0014, 0, $localCache, 0x01)
        Write-Log "SUCCESS: Wallpaper applied to desktop."
    } else {
        Write-Log "ERROR: All download attempts failed for $model. Please check GitHub filenames."
    }

} catch {
    Write-Log "Fatal Error: $($_.Exception.Message)"
}