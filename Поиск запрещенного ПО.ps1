# Добавляются .NET-библиотеки, необходимые для работы с графическим интерфейсом
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Глобальные переменные
$global:CancelOperation = $false
$global:ReportData = $null
$global:DomainName = ""

# --- Основные функции ---
# Получение списка установленного программного обеспечения
function Get-InstalledSoftware {
    param (
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        if ($global:CancelOperation) { return @() }

        $software = @()
        $scriptBlock = {
            $paths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            
            foreach ($path in $paths) {
                $items = Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
                foreach ($item in $items) {
                    [PSCustomObject]@{
                        Name = $item.DisplayName
                        Version = $item.DisplayVersion
                        Publisher = $item.Publisher
                        Computer = $env:COMPUTERNAME
                    }
                }
            }
        }

        if ($Credential) {
            $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential -ErrorAction Stop
        }
        else {
            $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
        }

        foreach ($item in $results) {
            $software += $item
        }
        
        return $software
    }
    catch {
        Write-Warning "Ошибка при проверке компьютера $ComputerName : $_"
        return @()
    }
}
# Загрузка списка запрещенных программ
function Import-RestrictedSoftwareList {
    param (
        [string]$Path
    )
    try {
        $list = Import-Csv -Path $Path | Select-Object -ExpandProperty 'Название программы'
        return $list
    }
    catch {
        [System.Windows.MessageBox]::Show("Ошибка при загрузке списка запрещенного ПО: $_", "Ошибка", "OK", "Error")
        return @()
    }
}
# Поиск запрещенных программ в предоставленном списке
function Find-RestrictedSoftware {
    param (
        [array]$InstalledSoftware,
        [array]$RestrictedList
    )
    $found = @()
    foreach ($app in $InstalledSoftware) {
        foreach ($restrictedName in $RestrictedList) {
            if ($app.Name -like "*$restrictedName*") {
                $found += [PSCustomObject]@{
                    Компьютер = $app.Computer
                    Название = $app.Name
                    Версия = $app.Version
                    Издатель = $app.Publisher
                    ЗапрещенноеПО = $restrictedName
                }
            }
        }
    }
    return $found
}
# Получение списков ПЭВМ домена через подключение к AD
function Get-DomainComputers {
    param (
        [string]$Domain,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $searcher = New-Object DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectCategory=computer)(objectClass=computer))"
        $searcher.SearchRoot = "LDAP://$Domain"
        
        if ($Credential) {
            $domainEntry = New-Object DirectoryServices.DirectoryEntry("LDAP://$Domain", $Credential.UserName, $Credential.GetNetworkCredential().Password)
            $searcher.SearchRoot = $domainEntry
        }
        
        $computers = $searcher.FindAll() | ForEach-Object { $_.Properties.name }
        return $computers
    }
    catch {
        [System.Windows.MessageBox]::Show("Ошибка при получении списка компьютеров из AD: $_", "Ошибка", "OK", "Error")
        return @()
    }
}
# Получение списков пользователей на ПЭВМ
function Get-ComputerUsers {
    param (
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $scriptBlock = {
            $users = @()
            $sessions = quser | Select-Object -Skip 1
            
            foreach ($session in $sessions) {
                if ($session -match "(\S+)\s+") {
                    $users += $matches[1]
                }
            }
            
            $users | Select-Object -Unique
        }

        if ($Credential) {
            $users = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential -ErrorAction Stop
        }
        else {
            $users = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
        }

        return $users
    }
    catch {
        Write-Warning "Ошибка при получении пользователей компьютера $ComputerName : $_"
        return @()
    }
}
#Поиск электронной почты в user AD
function Get-UserEmailFromAD {
    param (
        [string]$Username,
        [string]$Domain,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $searcher = New-Object DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectCategory=user)(samaccountname=$Username))"
        $searcher.SearchRoot = "LDAP://$Domain"
        
        if ($Credential) {
            $domainEntry = New-Object DirectoryServices.DirectoryEntry("LDAP://$Domain", $Credential.UserName, $Credential.GetNetworkCredential().Password)
            $searcher.SearchRoot = $domainEntry
        }
        
        $result = $searcher.FindOne()
        
        if ($result -and $result.Properties.mail) {
            return $result.Properties.mail[0]
        }
        return $null
    }
    catch {
        Write-Warning "Ошибка при поиске почты пользователя $Username : $_"
        return $null
    }
}
#Формирование письма
function Send-EmailReport {
    param (
        [string]$To,
        [string]$Subject,
        [array]$RestrictedSoftware,
        [string]$SmtpServer,
        [System.Management.Automation.PSCredential]$EmailCredential
    )
    try {
        $from = "security@$($global:DomainName)"
        
        if ($RestrictedSoftware.Count -gt 0) {
            $html = @"
<html>
<head>
<style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; }
    tr:nth-child(even) { background-color: #f9f9f9; }
</style>
</head>
<body>
<h2>Отчет о запрещенном ПО</h2>
<p>На вашем компьютере обнаружено следующее запрещенное ПО:</p>
<table>
<tr><th>Компьютер</th><th>Программа</th><th>Версия</th><th>Издатель</th></tr>
"@
            foreach ($item in $RestrictedSoftware) {
                $html += "<tr><td>$($item.Компьютер)</td><td>$($item.Название)</td><td>$($item.Версия)</td><td>$($item.Издатель)</td></tr>"
            }
            $html += @"
</table>
<p>Пожалуйста, удалите указанное программное обеспечение.</p>
<p>С уважением,<br>Отдел информационной безопасности</p>
</body>
</html>
"@
        }
        else {
            $html = @"
<html>
<body>
<h2>Отчет о запрещенном ПО</h2>
<p>На вашем компьютере запрещенное ПО не обнаружено.</p>
<p>С уважением,<br>Отдел информационной безопасности</p>
</body>
</html>
"@
        }
        
        $mailParams = @{
            From = $from
            To = $To
            Subject = $Subject
            Body = $html
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            UseSsl = $true
            Port = 587
        }
        
        if ($EmailCredential) {
            $mailParams.Credential = $EmailCredential
        }
        
        Send-MailMessage @mailParams
        return $true
    }
    catch {
        Write-Warning "Ошибка при отправке письма: $_"
        return $false
    }
}
#Сохранение
function Save-Report {
    param (
        [array]$Data,
        [string]$Format,
        [string]$FilePath
    )
    try {
        $directory = [System.IO.Path]::GetDirectoryName($FilePath)
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        switch ($Format) {
            "CSV" {
                $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
                break
            }
            "JSON" {
                $Data | ConvertTo-Json -Depth 5 | Out-File -FilePath $FilePath -Encoding UTF8
                break
            }
            default {
                throw "Неподдерживаемый формат отчета: $Format"
            }
        }
        
        return $true
    }
    catch {
        Write-Warning "Ошибка при сохранении отчета: $_"
        return $false
    }
}

# --- Графический интерфейс ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Проверка запрещенного ПО в домене"
$form.Size = New-Object System.Drawing.Size(900, 750)
$form.StartPosition = "CenterScreen"
$form.Add_FormClosing({
    if (-not $global:CancelOperation -and $global:ReportData -ne $null) {
        $result = [System.Windows.MessageBox]::Show("Сохранить отчет перед закрытием?", "Подтверждение", "YesNoCancel", "Question")
        
        if ($result -eq "Yes") {
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "CSV файлы (*.csv)|*.csv|JSON файлы (*.json)|*.json"
            $saveDialog.Title = "Сохранить отчет"
            
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $format = if ($saveDialog.FileName.EndsWith(".json")) { "JSON" } else { "CSV" }
                Save-Report -Data $global:ReportData -Format $format -FilePath $saveDialog.FileName
            }
        }
        elseif ($result -eq "Cancel") {
            $_.Cancel = $true
        }
    }
})

# Настройки домена
$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Location = New-Object System.Drawing.Point(20, 20)
$labelDomain.Size = New-Object System.Drawing.Size(300, 20)
$labelDomain.Text = "Домен (например: local.corp):"

$textBoxDomain = New-Object System.Windows.Forms.TextBox
$textBoxDomain.Location = New-Object System.Drawing.Point(20, 40)
$textBoxDomain.Size = New-Object System.Drawing.Size(300, 20)
$textBoxDomain.Text = "local.corp"

# Учетные данные администратора
$labelAdminUser = New-Object System.Windows.Forms.Label
$labelAdminUser.Location = New-Object System.Drawing.Point(20, 70)
$labelAdminUser.Size = New-Object System.Drawing.Size(150, 20)
$labelAdminUser.Text = "Пользователь домена:"

$textBoxAdminUser = New-Object System.Windows.Forms.TextBox
$textBoxAdminUser.Location = New-Object System.Drawing.Point(180, 70)
$textBoxAdminUser.Size = New-Object System.Drawing.Size(200, 20)

$labelAdminPass = New-Object System.Windows.Forms.Label
$labelAdminPass.Location = New-Object System.Drawing.Point(20, 100)
$labelAdminPass.Size = New-Object System.Drawing.Size(150, 20)
$labelAdminPass.Text = "Пароль:"

$textBoxAdminPass = New-Object System.Windows.Forms.MaskedTextBox
$textBoxAdminPass.Location = New-Object System.Drawing.Point(180, 100)
$textBoxAdminPass.Size = New-Object System.Drawing.Size(200, 20)
$textBoxAdminPass.PasswordChar = '*'

# Выбор файла с запрещенным ПО
$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Location = New-Object System.Drawing.Point(20, 130)
$labelFile.Size = New-Object System.Drawing.Size(300, 20)
$labelFile.Text = "Файл со списком запрещенного ПО (CSV):"

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Location = New-Object System.Drawing.Point(20, 150)
$textBoxFile.Size = New-Object System.Drawing.Size(500, 20)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(530, 150)
$buttonBrowse.Size = New-Object System.Drawing.Size(75, 20)
$buttonBrowse.Text = "Обзор"

$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxFile.Text = $openFileDialog.FileName
    }
})

# Настройки отправки
$groupBoxEmail = New-Object System.Windows.Forms.GroupBox
$groupBoxEmail.Location = New-Object System.Drawing.Point(20, 180)
$groupBoxEmail.Size = New-Object System.Drawing.Size(850, 200)
$groupBoxEmail.Text = "Настройки отправки отчетов"

$radioAdminOnly = New-Object System.Windows.Forms.RadioButton
$radioAdminOnly.Location = New-Object System.Drawing.Point(20, 30)
$radioAdminOnly.Size = New-Object System.Drawing.Size(300, 20)
$radioAdminOnly.Text = "Отправить только системному администратору"
$radioAdminOnly.Checked = $true

$radioAllUsers = New-Object System.Windows.Forms.RadioButton
$radioAllUsers.Location = New-Object System.Drawing.Point(20, 60)
$radioAllUsers.Size = New-Object System.Drawing.Size(300, 20)
$radioAllUsers.Text = "Отправить всем пользователям"

$labelAdminEmail = New-Object System.Windows.Forms.Label
$labelAdminEmail.Location = New-Object System.Drawing.Point(20, 90)
$labelAdminEmail.Size = New-Object System.Drawing.Size(200, 20)
$labelAdminEmail.Text = "Email администратора:"

$textBoxAdminEmail = New-Object System.Windows.Forms.TextBox
$textBoxAdminEmail.Location = New-Object System.Drawing.Point(220, 90)
$textBoxAdminEmail.Size = New-Object System.Drawing.Size(300, 20)
$textBoxAdminEmail.Text = "admin@local.corp"

$labelSmtpServer = New-Object System.Windows.Forms.Label
$labelSmtpServer.Location = New-Object System.Drawing.Point(20, 120)
$labelSmtpServer.Size = New-Object System.Drawing.Size(200, 20)
$labelSmtpServer.Text = "SMTP сервер:"

$textBoxSmtpServer = New-Object System.Windows.Forms.TextBox
$textBoxSmtpServer.Location = New-Object System.Drawing.Point(220, 120)
$textBoxSmtpServer.Size = New-Object System.Drawing.Size(300, 20)
$textBoxSmtpServer.Text = "smtp.local.corp"

$labelEmailUser = New-Object System.Windows.Forms.Label
$labelEmailUser.Location = New-Object System.Drawing.Point(20, 150)
$labelEmailUser.Size = New-Object System.Drawing.Size(200, 20)
$labelEmailUser.Text = "Пользователь SMTP:"

$textBoxEmailUser = New-Object System.Windows.Forms.TextBox
$textBoxEmailUser.Location = New-Object System.Drawing.Point(220, 150)
$textBoxEmailUser.Size = New-Object System.Drawing.Size(200, 20)

$labelEmailPass = New-Object System.Windows.Forms.Label
$labelEmailPass.Location = New-Object System.Drawing.Point(430, 150)
$labelEmailPass.Size = New-Object System.Drawing.Size(100, 20)
$labelEmailPass.Text = "Пароль:"

$textBoxEmailPass = New-Object System.Windows.Forms.MaskedTextBox
$textBoxEmailPass.Location = New-Object System.Drawing.Point(530, 150)
$textBoxEmailPass.Size = New-Object System.Drawing.Size(200, 20)
$textBoxEmailPass.PasswordChar = '*'

$groupBoxEmail.Controls.AddRange(@(
    $radioAdminOnly,
    $radioAllUsers,
    $labelAdminEmail,
    $textBoxAdminEmail,
    $labelSmtpServer,
    $textBoxSmtpServer,
    $labelEmailUser,
    $textBoxEmailUser,
    $labelEmailPass,
    $textBoxEmailPass
))

# Список исключений
$labelExclude = New-Object System.Windows.Forms.Label
$labelExclude.Location = New-Object System.Drawing.Point(20, 390)
$labelExclude.Size = New-Object System.Drawing.Size(300, 20)
$labelExclude.Text = "Исключить пользователей (через запятую):"

$textBoxExclude = New-Object System.Windows.Forms.TextBox
$textBoxExclude.Location = New-Object System.Drawing.Point(20, 410)
$textBoxExclude.Size = New-Object System.Drawing.Size(400, 20)
$textBoxExclude.Text = "admin,guest,test"

# Формат отчета
$groupBoxFormat = New-Object System.Windows.Forms.GroupBox
$groupBoxFormat.Location = New-Object System.Drawing.Point(430, 390)
$groupBoxFormat.Size = New-Object System.Drawing.Size(200, 80)
$groupBoxFormat.Text = "Формат отчета"

$radioFormatCsv = New-Object System.Windows.Forms.RadioButton
$radioFormatCsv.Location = New-Object System.Drawing.Point(20, 20)
$radioFormatCsv.Size = New-Object System.Drawing.Size(50, 20)
$radioFormatCsv.Text = "CSV"
$radioFormatCsv.Checked = $true

$radioFormatJson = New-Object System.Windows.Forms.RadioButton
$radioFormatJson.Location = New-Object System.Drawing.Point(20, 50)
$radioFormatJson.Size = New-Object System.Drawing.Size(50, 20)
$radioFormatJson.Text = "JSON"

$groupBoxFormat.Controls.AddRange(@($radioFormatCsv, $radioFormatJson))

# Кнопки управления
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Location = New-Object System.Drawing.Point(20, 450)
$buttonRun.Size = New-Object System.Drawing.Size(150, 30)
$buttonRun.Text = "Запустить проверку"

$buttonStop = New-Object System.Windows.Forms.Button
$buttonStop.Location = New-Object System.Drawing.Point(180, 450)
$buttonStop.Size = New-Object System.Drawing.Size(150, 30)
$buttonStop.Text = "Остановить"
$buttonStop.Enabled = $false

$buttonSaveReport = New-Object System.Windows.Forms.Button
$buttonSaveReport.Location = New-Object System.Drawing.Point(340, 450)
$buttonSaveReport.Size = New-Object System.Drawing.Size(150, 30)
$buttonSaveReport.Text = "Сохранить отчет"
$buttonSaveReport.Enabled = $false

# Лог выполнения
$textBoxLog = New-Object System.Windows.Forms.TextBox
$textBoxLog.Location = New-Object System.Drawing.Point(20, 500)
$textBoxLog.Size = New-Object System.Drawing.Size(850, 200)
$textBoxLog.Multiline = $true
$textBoxLog.ScrollBars = "Vertical"
$textBoxLog.ReadOnly = $true

# Обработчики событий
$buttonRun.Add_Click({
    if (-not (Test-Path $textBoxFile.Text)) {
        [System.Windows.MessageBox]::Show("Укажите корректный файл со списком запрещенного ПО", "Ошибка", "OK", "Error")
        return
    }

    if ([string]::IsNullOrWhiteSpace($textBoxDomain.Text)) {
        [System.Windows.MessageBox]::Show("Укажите домен", "Ошибка", "OK", "Error")
        return
    }

    $global:CancelOperation = $false
    $global:ReportData = $null
    $global:DomainName = $textBoxDomain.Text
    $buttonRun.Enabled = $false
    $buttonStop.Enabled = $true
    $buttonSaveReport.Enabled = $false
    $textBoxLog.Clear()
    $textBoxLog.AppendText("Начало проверки...`r`n")
    
    # Создаем credential для домена
    $domainCred = $null
    if (-not [string]::IsNullOrWhiteSpace($textBoxAdminUser.Text)) {
        $securePass = ConvertTo-SecureString $textBoxAdminPass.Text -AsPlainText -Force
        $domainCred = New-Object System.Management.Automation.PSCredential ($textBoxAdminUser.Text, $securePass)
    }
    
    # Создаем credential для почты
    $emailCred = $null
    if (-not [string]::IsNullOrWhiteSpace($textBoxEmailUser.Text)) {
        $secureEmailPass = ConvertTo-SecureString $textBoxEmailPass.Text -AsPlainText -Force
        $emailCred = New-Object System.Management.Automation.PSCredential ($textBoxEmailUser.Text, $secureEmailPass)
    }
    
    # Получаем список запрещенного ПО
    $restrictedList = Import-RestrictedSoftwareList -Path $textBoxFile.Text
    if ($restrictedList.Count -eq 0) {
        $textBoxLog.AppendText("Не удалось загрузить список запрещенного ПО`r`n")
        $buttonRun.Enabled = $true
        $buttonStop.Enabled = $false
        return
    }
    
    $textBoxLog.AppendText("Загружен список запрещенного ПО ($($restrictedList.Count) позиций)`r`n")
    
    # Получаем список компьютеров в домене
    $ldapDomain = "DC=" + $textBoxDomain.Text.Replace(".", ",DC=")
    $computers = Get-DomainComputers -Domain $ldapDomain -Credential $domainCred
    if ($computers.Count -eq 0) {
        $textBoxLog.AppendText("Не удалось получить список компьютеров из AD`r`n")
        $buttonRun.Enabled = $true
        $buttonStop.Enabled = $false
        return
    }
    
    $textBoxLog.AppendText("Найдено компьютеров в домене: $($computers.Count)`r`n")
    $textBoxLog.AppendText("Начало сканирования...`r`n")
    
    # Проверяем каждый компьютер
    $allRestrictedSoftware = @()
    $userReports = @{}
    $processedComputers = 0
    
    foreach ($computer in $computers) {
        if ($global:CancelOperation) {
            $textBoxLog.AppendText("Проверка остановлена пользователем`r`n")
            break
        }
        
        $processedComputers++
        $textBoxLog.AppendText("Проверка компьютера $processedComputers из $($computers.Count): $computer... ")
        
        $software = Get-InstalledSoftware -ComputerName $computer -Credential $domainCred
        $users = Get-ComputerUsers -ComputerName $computer -Credential $domainCred
        
        $found = Find-RestrictedSoftware -InstalledSoftware $software -RestrictedList $restrictedList
        
        if ($found.Count -gt 0) {
            $textBoxLog.AppendText("найдено $($found.Count) запрещенных программ`r`n")
            $allRestrictedSoftware += $found
            
            # Добавляем информацию о пользователях
            foreach ($user in $users) {
                if (-not $userReports.ContainsKey($user)) {
                    $userReports[$user] = @()
                }
                $userReports[$user] += $found
            }
        }
        else {
            $textBoxLog.AppendText("запрещенное ПО не найдено`r`n")
        }
    }
    
    # Формируем отчет
    $global:ReportData = $allRestrictedSoftware
    $textBoxLog.AppendText("`r`nПроверка завершена. Найдено всего: $($allRestrictedSoftware.Count) запрещенных программ`r`n")
    
    # Отправка отчетов
    if (-not $global:CancelOperation) {
        if ($radioAdminOnly.Checked) {
            $textBoxLog.AppendText("Отправка отчета администратору... ")
            $result = Send-EmailReport -To $textBoxAdminEmail.Text -Subject "Отчет о запрещенном ПО" -RestrictedSoftware $allRestrictedSoftware -SmtpServer $textBoxSmtpServer.Text -EmailCredential $emailCred
            if ($result) {
                $textBoxLog.AppendText("успешно`r`n")
            }
            else {
                $textBoxLog.AppendText("ошибка`r`n")
            }
        }
        else {
            # Отправка всем пользователям
            $excludeUsers = $textBoxExclude.Text -split ',' | ForEach-Object { $_.Trim() }
            
            $textBoxLog.AppendText("`r`nПодготовка отчетов для пользователей...`r`n")
            
            $totalSent = 0
            $totalFailed = 0
            
            foreach ($user in $userReports.Keys) {
                if ($global:CancelOperation) {
                    $textBoxLog.AppendText("Отправка прервана пользователем`r`n")
                    break
                }
                
                if ($excludeUsers -contains $user) {
                    $textBoxLog.AppendText("Пропускаем пользователя $user (в списке исключений)`r`n")
                    continue
                }
                
                $email = Get-UserEmailFromAD -Username $user -Domain $ldapDomain -Credential $domainCred
                
                if ($email) {
                    $textBoxLog.AppendText("Отправка отчета пользователю $user ($email)... ")
                    
                    $userSoftware = $userReports[$user] | Where-Object { $_.Пользователь -eq $user }
                    $result = Send-EmailReport -To $email -Subject "Уведомление о запрещенном ПО" -RestrictedSoftware $userSoftware -SmtpServer $textBoxSmtpServer.Text -EmailCredential $emailCred
                    
                    if ($result) {
                        $textBoxLog.AppendText("успешно`r`n")
                        $totalSent++
                    }
                    else {
                        $textBoxLog.AppendText("ошибка`r`n")
                        $totalFailed++
                    }
                }
                else {
                    $textBoxLog.AppendText("Не удалось найти email для пользователя $user`r`n")
                    $totalFailed++
                }
            }
            
            $textBoxLog.AppendText("`r`nИтоги отправки:`r`n")
            $textBoxLog.AppendText("Успешно отправлено: $totalSent`r`n")
            $textBoxLog.AppendText("Не удалось отправить: $totalFailed`r`n")
        }
    }
    
    $buttonRun.Enabled = $true
    $buttonStop.Enabled = $false
    $buttonSaveReport.Enabled = $true
    
    if (-not $global:CancelOperation) {
        $saveReport = [System.Windows.MessageBox]::Show("Проверка завершена. Сохранить отчет?", "Отчет", "YesNo", "Question")
        if ($saveReport -eq "Yes") {
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "CSV файлы (*.csv)|*.csv|JSON файлы (*.json)|*.json"
            $saveDialog.Title = "Сохранить отчет"
            
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $format = if ($saveDialog.FileName.EndsWith(".json")) { "JSON" } else { "CSV" }
                $result = Save-Report -Data $global:ReportData -Format $format -FilePath $saveDialog.FileName
                
                if ($result) {
                    $textBoxLog.AppendText("Отчет успешно сохранен: $($saveDialog.FileName)`r`n")
                }
                else {
                    $textBoxLog.AppendText("Ошибка при сохранении отчета`r`n")
                }
            }
        }
    }
})

$buttonStop.Add_Click({
    $global:CancelOperation = $true
    $buttonStop.Enabled = $false
    $textBoxLog.AppendText("Запрошена остановка проверки...`r`n")
})

$buttonSaveReport.Add_Click({
    if ($global:ReportData -eq $null) {
        [System.Windows.MessageBox]::Show("Нет данных для отчета", "Ошибка", "OK", "Error")
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV файлы (*.csv)|*.csv|JSON файлы (*.json)|*.json"
    $saveDialog.Title = "Сохранить отчет"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $format = if ($saveDialog.FileName.EndsWith(".json")) { "JSON" } else { "CSV" }
        $result = Save-Report -Data $global:ReportData -Format $format -FilePath $saveDialog.FileName
        
        if ($result) {
            $textBoxLog.AppendText("Отчет успешно сохранен: $($saveDialog.FileName)`r`n")
        }
        else {
            $textBoxLog.AppendText("Ошибка при сохранении отчета`r`n")
        }
    }
})

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(500, 450)
$buttonCancel.Size = New-Object System.Drawing.Size(150, 30)
$buttonCancel.Text = "Закрыть"
$buttonCancel.Add_Click({
    $form.Close()
})

# Добавляем элементы на форму
$form.Controls.AddRange(@(
    $labelDomain,
    $textBoxDomain,
    $labelAdminUser,
    $textBoxAdminUser,
    $labelAdminPass,
    $textBoxAdminPass,
    $labelFile,
    $textBoxFile,
    $buttonBrowse,
    $groupBoxEmail,
    $labelExclude,
    $textBoxExclude,
    $groupBoxFormat,
    $buttonRun,
    $buttonStop,
    $buttonSaveReport,
    $buttonCancel,
    $textBoxLog
))

# Запуск формы
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)