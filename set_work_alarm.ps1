# 設定 PowerShell 允許執行腳本（避免權限問題）
Set-ExecutionPolicy Bypass -Scope Process -Force

$filePath = "D:\PoweShell\loginTimeList.json"
$loginTime = Get-Date
$formattedTime = $loginTime.ToString()

# 檢查檔案是否存在
if (Test-Path $filePath) {
    # 讀取現有的 JSON 檔案
    try {
        $jsonContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
        
        # 確保內容是陣列
        if ($null -eq $jsonContent) {
            $jsonContent = @()
        }
        elseif ($jsonContent -isnot [Array]) {
            $jsonContent = @($jsonContent)
        }
    }
    catch {
        # If the file is not in a valid JSON format, reinitialize to an empty array
        Write-Host "This is not a valid JSON file."
        # $jsonContent = @()
    }
}
else {
    # 如果檔案不存在，初始化為空陣列
    $jsonContent = @()
}

# 新增當前的登入時間
$newEntry = @{loginTime = $formattedTime}
$jsonContent += $newEntry

# 將結果轉換為 JSON 並寫入檔案
$jsonContent | ConvertTo-Json -Depth 10 | Out-File -FilePath $filePath -Encoding UTF8

$loginData = Get-Content "D:\PoweShell\loginTimeList.json" | ConvertFrom-Json
$earliestLogin = ($loginData | Sort-Object loginTime | Select-Object -First 1).loginTime
$alarmTime = (Get-Date $earliestLogin).AddHours(9.5)

if ($alarmTime -lt $loginTime ) {
    $toast = New-Object -ComObject WScript.Shell
    $toast.Popup("The alarm time is earlier than the login time.", 16, "Eorror", 64)
    Unregister-ScheduledTask -TaskName "WorkAlarm" -Confirm:$false
    exit
}

# 使用 Windows Toast Notification
$toast = New-Object -ComObject WScript.Shell
$toast.Popup("Today's first login time : $($earliestLogin.ToString())`n Alarm time : $($alarmTime.ToString())", 16, "Notification", 64)

# 建立執行動作
$actionScript = @'
$toast = New-Object -ComObject WScript.Shell
$toast.Popup("TIME TO HEAD HOME !!", 32, "Notification", 64)

# 執行完通知後刪除此排程任務
Unregister-ScheduledTask -TaskName "WorkAlarm" -Confirm:$false
'@

# 將腳本保存到臨時檔案
$tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
$actionScript | Out-File -FilePath $tempFile

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$tempFile`""

$trigger = New-ScheduledTaskTrigger -Once -At $alarmTime

# 註冊排程工作
Register-ScheduledTask -TaskName "WorkAlarm" -Trigger $trigger -Action $action -Description "Work Alarm" -User $env:UserName -RunLevel Limited
