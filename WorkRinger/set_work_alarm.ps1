# 設定 PowerShell 允許執行腳本（避免權限問題）
# 這行命令用於設定 PowerShell 的執行策略，允許在當前進程中執行腳本，避免權限問題。
Set-ExecutionPolicy Bypass -Scope Process -Force

# 獲取當前登入時間
$loginTime = Get-Date
# 最早登入時間
$earliestLoginTime = Get-Date -Hour 08 -Minute 30 -Second 0 -Millisecond 0

if ($loginTime -lt $earliestLoginTime) {
    $loginTime = $earliestLoginTime
}

# 將登入時間轉換為字串格式
$formattedTime = $loginTime.ToString()
# 指定 JSON 檔案的路徑，用于存儲登入時間列表
$saveTimeListPath = "D:\PoweShell\WorkRinger\loginTimeList.json" 
# 定義排程任務的名稱
$taskName = "End-of-Work-Bell"

# 檢查 JSON 檔案是否存在
if (Test-Path $saveTimeListPath) {
    
    try {
        # 讀取 JSON 檔案的內容，轉換為 PowerShell 物件
        $jsonContent = Get-Content -Path $saveTimeListPath -Raw | ConvertFrom-Json
        
        # 如果讀取到的內容為空，則初始化為空陣列
        if ($null -eq $jsonContent) {
            $jsonContent = @()
        }
        # 如果讀取到的內容不是陣列，則將其轉換為陣列
        elseif ($jsonContent -isnot [Array]) {
            $jsonContent = @($jsonContent)
        }
    }
    catch {
        # 如果讀取 JSON 檔案時發生錯誤，則顯示錯誤訊息
        $toast = New-Object -ComObject WScript.Shell
        $toast.Popup("This is not a valid JSON file.", 16, "Error", 64)
        # 移除已經註冊的排程任務
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        # 結束腳本
        exit
    }
}
else {
    # 如果 JSON 檔案路徑無效，則顯示錯誤訊息
    $toast = New-Object -ComObject WScript.Shell
    $toast.Popup("Your saveTimeListPath is not valid.", 16, "Error", 64)
    # 移除已經註冊的排程任務
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    # 結束腳本
    exit
}

# 新增當前的登入時間到登入時間列表
$newLoginTime = @{loginTime = $formattedTime }
$jsonContent += $newLoginTime

# 將結果轉換為 JSON 並寫入檔案
$jsonContent | ConvertTo-Json -Depth 10 | Out-File -FilePath $saveTimeListPath -Encoding UTF8

# 讀取 JSON 檔案，獲取登入時間列表
$loginData = Get-Content $saveTimeListPath | ConvertFrom-Json
# 尋找最早的登入時間
$earliestLogin = ($loginData | Sort-Object loginTime | Select-Object -First 1).loginTime
# 設定工作時間（小時）
$workTime = 9.5
# 計算警報時間
$alarmTime = (Get-Date $earliestLogin).AddHours($workTime)

# 檢查警報時間是否早於登入時間
if ($alarmTime -lt $loginTime ) {
    # 如果警報時間早於登入時間，則顯示錯誤訊息
    $toast = New-Object -ComObject WScript.Shell
    $toast.Popup("The alarm time is earlier than the login time.", 16, "Error", 64)
    # 移除已經註冊的排程任務
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    # 結束腳本
    exit
}

# 使用 Windows Toast Notification 顯示訊息
$toast = New-Object -ComObject WScript.Shell
$toast.Popup("Today's first login time : $($earliestLogin.ToString())`n Alarm time : $($alarmTime.ToString())", 16, "Notification", 64)

# 定義執行動作腳本
$actionScript = @'
# 清空 JSON 檔案內容，但保留檔案
if(-not ($saveTimeListPath)) {
    $saveTimeListPath = "D:\PoweShell\WorkRinger\loginTimeList.json"
}

$null | ConvertTo-Json | Out-File -FilePath $saveTimeListPath -Encoding UTF8

# 顯示通知
$toast = New-Object -ComObject WScript.Shell
$toast.Popup("TIME TO HEAD HOME !!", 32, "Notification", 64)

# 取消排程任務
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
'@

# 將腳本保存到臨時檔案
$tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
$actionScript | Out-File -FilePath $tempFile

# 定義執行動作
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$tempFile`""

# 定義觸發器，設定警報時間
$trigger = New-ScheduledTaskTrigger -Once -At $alarmTime

# 註冊排程工作
Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Description "Work Ringer" -User $env:UserName -RunLevel Limited