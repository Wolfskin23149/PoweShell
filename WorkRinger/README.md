這個手冊將幫助您設置一個自動化工具，記錄登入時間並在工作 9.5 小時後提醒您下班。請按照以下步驟操作。

### 步驟 1：建立並儲存 PowerShell 腳本

1. 打開記事本 
   - 按 Windows 鍵，輸入 記事本，按 Enter。

2. 貼上腳本內容 
   - 複製以下腳本並貼到記事本中：

   ```powershell
   # 設定 PowerShell 允許執行腳本（避免權限問題）
   # 這行命令用於設定 PowerShell 的執行策略，允許在當前進程中執行腳本，避免權限問題。
   Set-ExecutionPolicy Bypass -Scope Process -Force

   # 獲取當前登入時間
   $loginTime = Get-Date
   # 將登入時間轉換為字串格式
   $formattedTime = $loginTime.ToString()
   # 指定 JSON 檔案的路徑，用于存儲登入時間列表
   $saveTimeListPath = "D:\PoweShell\loginTimeList.json" 
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
           $toast.Popup("This is not a valid JSON file.", 16, "Eorror", 64)
           # 移除已經註冊的排程任務
           Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
           # 結束腳本
           exit
       }
   }
   else {
       # 如果 JSON 檔案路徑無效，則顯示錯誤訊息
       $toast = New-Object -ComObject WScript.Shell
       $toast.Popup("Your saveTimeListPath is not valid.", 16, "Eorror", 64)
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
       $toast.Popup("The alarm time is earlier than the login time.", 16, "Eorror", 64)
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
   $toast = New-Object -ComObject WScript.Shell
   $toast.Popup("TIME TO HEAD HOME !!", 32, "Notification", 64)
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
   ```

3. 儲存檔案 
   - 點擊 檔案 → 另存新檔。
   - 檔案名稱：`WorkEndReminder.ps1`。
   - 儲存位置：`D:\PowerShell`（若無此資料夾，請先建立）。
   - 儲存類型：`所有檔案 (.)`。
   - 編碼：`UTF-8`。
   - 按 儲存。

### 步驟 2：設定工作排程器自動觸發

1. 打開工作排程器 
   - 按 Windows 鍵，輸入 工作排程器，按 Enter。

2. 建立新任務 
   - 在右側點擊 建立任務（不是「建立基本任務」）。

3. 設定任務名稱與權限 
   - 在「一般」標籤中：
     - 名稱：輸入 LoginTrigger。
     - 勾選「不論使用者是否登入都要執行」。
     - 勾選「使用最高權限執行」。

4. 設定觸發條件 
   - 切換到「觸發條件」標籤，點擊 新增。
   - 在「開始任務」下拉選單選擇 工作站解除鎖定時。
   - 按 確定。

5. 設定執行動作 
   - 切換到「動作」標籤，點擊 新增。
   - 動作：選擇 啟動程式。
   - 程式或指令碼：輸入 powershell.exe。
   - 加入引數：輸入 -ExecutionPolicy Bypass -File "D:\PowerShell\WorkEndReminder.ps1"。
   - 按 確定。

6. 儲存任務 
   - 點擊 確定 保存任務。
   - 如果跳出帳號密碼視窗，輸入您的 Windows 使用者名稱和密碼，然後按 確定。

### 步驟 3：測試與使用

1. 測試功能 
   - 鎖定螢幕（按 Windows 鍵 + L），然後解鎖。
   - 您應該會看到一個彈窗，顯示當天第一次登入時間和預計下班時間。

2. 每天使用 
   - 每次解鎖電腦時，腳本會自動記錄時間並設定下班提醒。

3. 到達下班時間（最早登入後 9.5 小時），會彈出「TIME TO HEAD HOME !!」提醒。

### 常見問題與解決方法

1. 沒有彈出提醒？ 
   - 確認 D:\PowerShell 資料夾和 WorkEndReminder.ps1 檔案存在。
   - 檢查工作排程器中 LoginTrigger 任務是否啟用。

2. 想調整工作時間？ 
   - 打開 WorkEndReminder.ps1，找到 $workTime = 9.5，改成您想要的小時數（例如 8），儲存後重新解鎖電腦測試。

### 注意事項

- 確保電腦保持開機，否則提醒無法觸發。
- 下班提醒後，當天任務會自動清除，次日解鎖時重新開始。
