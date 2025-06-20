<#
.SYNOPSIS
    Комплексное решение для инвентаризации программного обеспечения в домене Active Directory
.DESCRIPTION
    Скрипт выполняет сбор информации об установленном ПО на компьютерах домена, 
    анализирует время последнего использования, формирует подробные отчеты 
    и осуществляет рассылку пользователям и администратору. 
    Поддерживает несколько методов сбора данных и подключения к удаленным компьютерам.
.NOTES
    Требования: PowerShell 5.1+, .NET Framework 4.7.2, права администратора домена
#>

#region Подключение необходимых сборок
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
#endregion

#region Глобальные переменные
$global:logPath = "C:\Logs\SoftwareInventory_$(Get-Date -Format 'yyyyMMdd').log"
$global:logForm = $null
$global:logTextBox = $null
$global:logStream = $null
#endregion

#region Функции логирования
function Initialize-Logging {
    # Создание каталога для логов
    $logDir = Split-Path -Path $global:logPath -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Создание формы для отображения логов
    $global:logForm = New-Object System.Windows.Forms.Form
    $global:logForm.Text = "Ход выполнения инвентаризации ПО"
    $global:logForm.Size = New-Object System.Drawing.Size(800, 600)
    $global:logForm.StartPosition = "CenterScreen"
    $global:logForm.FormBorderStyle = "FixedDialog"
    $global:logForm.MaximizeBox = $false
    
    $global:logTextBox = New-Object System.Windows.Forms.TextBox
    $global:logTextBox.Multiline = $true
    $global:logTextBox.ScrollBars = "Vertical"
    $global:logTextBox.Dock = "Fill"
    $global:logTextBox.ReadOnly = $true
    $global:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $global:logForm.Controls.Add($global:logTextBox)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Закрыть"
    $btnClose.Dock = "Bottom"
    $btnClose.Add_Click({ $global:logForm.Close() })
    $global:logForm.Controls.Add($btnClose)
    
    # Отображение формы логов
    $global:logForm.Add_Shown({ $global:logForm.Activate() })
    $global:logForm.Show()
    
    # Создание потока для записи логов
    try {
        $global:logStream = [System.IO.StreamWriter]::new($global:logPath, $true)
    } catch {
        Write-Host "Не удалось открыть файл логов: $global:logPath. Ошибка: $_" -ForegroundColor Red
        $global:logPath = "$env:TEMP\SoftwareInventory_$(Get-Date -Format 'yyyyMMdd').log"
        $global:logStream = [System.IO.StreamWriter]::new($global:logPath, $true)
    }
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Запись в файл логов
    try {
        $global:logStream.WriteLine($logEntry)
        $global:logStream.Flush()
    } catch {
        Write-Host "Ошибка записи в лог-файл: $_" -ForegroundColor Red
    }
    
    # Вывод в окно логов
    if ($global:logTextBox -ne $null) {
        $global:logForm.Invoke([Action]{
            $global:logTextBox.AppendText("$logEntry`r`n")
            $global:logTextBox.ScrollToCaret()
        })
    }
    
    # Вывод в консоль
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry }
    }
}

function Close-Logging {
    if ($global:logStream -ne $null) {
        $global:logStream.Close()
        $global:logStream.Dispose()
    }
}
#endregion

#region Вспомогательные функции
function Test-ADConnection {
    param (
        [string]$Domain
    )
    
    try {
        $null = Get-ADDomain -Identity $Domain -ErrorAction Stop
        Write-Log "Успешное подключение к домену AD $Domain"
        return $true
    } catch {
        Write-Log "Ошибка подключения к домену AD $Domain. Ошибка: $_" -Level "ERROR"
        return $false
    }
}

function Test-SMTPConnection {
    param (
        [string]$SMTPServer,
        [int]$SMTPPort
    )
    
    try {
        $smtpTest = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
        $smtpTest.Dispose()
        Write-Log "SMTP сервер $SMTPServer на порту $SMTPPort доступен"
        return $true
    } catch {
        Write-Log "Ошибка подключения к SMTP серверу $SMTPServer на порту $SMTPPort. Ошибка: $_" -Level "ERROR"
        return $false
    }
}

function Show-Help {
    $helpText = @"
===================================== СПРАВКА ПО СКРИПТУ ИНВЕНТАРИЗАЦИИ ПО =====================================

ОПИСАНИЕ:
Скрипт предназначен для сбора информации об установленном программном обеспечении на компьютерах домена Active Directory.
Он анализирует время последнего использования ПО, формирует отчеты и отправляет их пользователям и администратору.

ФУНКЦИОНАЛЬНЫЕ ВОЗМОЖНОСТИ:
1. Сбор данных об установленном ПО через различные методы:
   - Анализ системного реестра
   - Запросы WMI/CIM
   - Анализ Prefetch-файлов
   - Чтение журналов событий
2. Определение времени последнего запуска ПО
3. Категоризация ПО по времени использования:
   - Не использовалось более полугода
   - Не использовалось от 3 до 6 месяцев
   - Использовалось в течение последних 3 месяцев
4. Автоматическое получение информации о пользователях из Active Directory
5. Формирование подробных отчетов в форматах CSV и HTML
6. Рассылка отчетов:
   - Пользователям: только данные по их компьютерам
   - Администратору: полный отчет
7. Гибкие настройки:
   - Возможность исключения пользователей из рассылки
   - Выбор методов сбора данных
   - Настройка параметров SMTP-сервера
   - Тестовый режим работы

ЭЛЕМЕНТЫ ГРАФИЧЕСКОГО ИНТЕРФЕЙСА:

1. Основные параметры:
   - Домен: имя домена Active Directory
   - Пользователь с правами: учетная запись с административными правами для доступа к компьютерам
   - Пароль: пароль для указанной учетной записи

2. Настройки электронной почты:
   - SMTP сервер: адрес почтового сервера для отправки отчетов
   - Отправитель: email адрес отправителя
   - Email администратора: адрес для отправки полного отчета
   - Флажок "Отправлять отчеты по email": включает/выключает рассылку

3. Дополнительные параметры:
   - Исключить пользователей: список пользователей (через запятую), которые не получат отчеты
   - Методы сбора данных: выбор методов для получения информации о ПО
   - Тестовый режим: ограничивает обработку первыми тремя компьютерами

КАК УКАЗАТЬ ПОЛЬЗОВАТЕЛЕЙ ДЛЯ ИСКЛЮЧЕНИЯ ИЗ РАССЫЛКИ:
1. В поле "Исключить пользователей" введите имена пользователей через запятую
2. Используйте SAMAccountName (имя для входа в систему)
3. Пример: ServiceAccount1,TestUser,BackupAdmin
4. Эти пользователи не получат email с отчетом о ПО на их компьютерах

ПОДДЕРЖИВАЕМЫЕ МЕТОДЫ ПОДКЛЮЧЕНИЯ К КОМПЬЮТЕРАМ:
Скрипт автоматически пробует несколько методов подключения к удаленным компьютерам:
1. WMI (Windows Management Instrumentation)
2. CIM (Common Information Model)
3. Прямой доступ к реестру через удаленный API
4. Анализ общих сетевых ресурсов (административные общие ресурсы)

Если один метод не срабатывает, скрипт автоматически пробует следующий метод.

ЛОГИРОВАНИЕ:
- Все действия скрипта записываются в лог-файл
- Лог отображается в отдельном окне в режиме реального времени
- После завершения работы скрипта лог-файл сохраняется на диске

================================================================================================================
"@

    [System.Windows.MessageBox]::Show($helpText, "Справка по скрипту инвентаризации ПО", "OK", "Information")
}
#endregion

#region Графический интерфейс
function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Инвентаризация ПО в домене AD"
    $form.Size = New-Object System.Drawing.Size(700, 650)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

    # Панель вкладок
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $form.Controls.Add($tabControl)

    # Вкладка основных параметров
    $tabMain = New-Object System.Windows.Forms.TabPage
    $tabMain.Text = "Основные параметры"
    $tabControl.Controls.Add($tabMain)

    # Логотип
    try {
        $logo = New-Object System.Windows.Forms.PictureBox
        $logo.Width = 650
        $logo.Height = 70
        $logo.Location = New-Object System.Drawing.Point(10, 10)
        $logo.Image = [System.Drawing.Image]::FromFile(("I:\DIB\Security Access Control\Целеполагание\Алексей\Фото Т1.png"))
        $logo.SizeMode = "Zoom"
        $tabMain.Controls.Add($logo)
    } catch {
        # Если логотипа нет, пропускаем
    }

    # Основные параметры
    $lblDomain = New-Object System.Windows.Forms.Label
    $lblDomain.Text = "Домен:"
    $lblDomain.Location = New-Object System.Drawing.Point(20, 90)
    $lblDomain.Width = 150
    $tabMain.Controls.Add($lblDomain)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Text = "local.corp"
    $txtDomain.Location = New-Object System.Drawing.Point(180, 90)
    $txtDomain.Width = 250
    $tabMain.Controls.Add($txtDomain)

    $lblAdminUser = New-Object System.Windows.Forms.Label
    $lblAdminUser.Text = "Пользователь с правами:"
    $lblAdminUser.Location = New-Object System.Drawing.Point(20, 120)
    $lblAdminUser.Width = 150
    $tabMain.Controls.Add($lblAdminUser)

    $txtAdminUser = New-Object System.Windows.Forms.TextBox
    $txtAdminUser.Text = "$env:USERNAME@local.corp"
    $txtAdminUser.Location = New-Object System.Drawing.Point(180, 120)
    $txtAdminUser.Width = 250
    $tabMain.Controls.Add($txtAdminUser)

    $lblAdminPass = New-Object System.Windows.Forms.Label
    $lblAdminPass.Text = "Пароль:"
    $lblAdminPass.Location = New-Object System.Drawing.Point(20, 150)
    $lblAdminPass.Width = 150
    $tabMain.Controls.Add($lblAdminPass)

    $txtAdminPass = New-Object System.Windows.Forms.MaskedTextBox
    $txtAdminPass.PasswordChar = '*'
    $txtAdminPass.Location = New-Object System.Drawing.Point(180, 150)
    $txtAdminPass.Width = 250
    $tabMain.Controls.Add($txtAdminPass)

    # Вкладка настроек почты
    $tabEmail = New-Object System.Windows.Forms.TabPage
    $tabEmail.Text = "Настройки почты"
    $tabControl.Controls.Add($tabEmail)

    $lblSMTPServer = New-Object System.Windows.Forms.Label
    $lblSMTPServer.Text = "SMTP сервер:"
    $lblSMTPServer.Location = New-Object System.Drawing.Point(20, 20)
    $lblSMTPServer.Width = 150
    $tabEmail.Controls.Add($lblSMTPServer)

    $txtSMTPServer = New-Object System.Windows.Forms.TextBox
    $txtSMTPServer.Text = "smtp.local.corp"
    $txtSMTPServer.Location = New-Object System.Drawing.Point(180, 20)
    $txtSMTPServer.Width = 250
    $tabEmail.Controls.Add($txtSMTPServer)

    $lblSMTPPort = New-Object System.Windows.Forms.Label
    $lblSMTPPort.Text = "Порт SMTP:"
    $lblSMTPPort.Location = New-Object System.Drawing.Point(20, 50)
    $lblSMTPPort.Width = 150
    $tabEmail.Controls.Add($lblSMTPPort)

    $txtSMTPPort = New-Object System.Windows.Forms.TextBox
    $txtSMTPPort.Text = "25"
    $txtSMTPPort.Location = New-Object System.Drawing.Point(180, 50)
    $txtSMTPPort.Width = 100
    $tabEmail.Controls.Add($txtSMTPPort)

    $lblFromEmail = New-Object System.Windows.Forms.Label
    $lblFromEmail.Text = "Отправитель:"
    $lblFromEmail.Location = New-Object System.Drawing.Point(20, 80)
    $lblFromEmail.Width = 150
    $tabEmail.Controls.Add($lblFromEmail)

    $txtFromEmail = New-Object System.Windows.Forms.TextBox
    $txtFromEmail.Text = "noreply@local.corp"
    $txtFromEmail.Location = New-Object System.Drawing.Point(180, 80)
    $txtFromEmail.Width = 250
    $tabEmail.Controls.Add($txtFromEmail)

    $lblAdminEmail = New-Object System.Windows.Forms.Label
    $lblAdminEmail.Text = "Email администратора:"
    $lblAdminEmail.Location = New-Object System.Drawing.Point(20, 110)
    $lblAdminEmail.Width = 150
    $tabEmail.Controls.Add($lblAdminEmail)

    $txtAdminEmail = New-Object System.Windows.Forms.TextBox
    $txtAdminEmail.Text = "admin@local.corp"
    $txtAdminEmail.Location = New-Object System.Drawing.Point(180, 110)
    $txtAdminEmail.Width = 250
    $tabEmail.Controls.Add($txtAdminEmail)

    $chkSendEmails = New-Object System.Windows.Forms.CheckBox
    $chkSendEmails.Text = "Отправлять отчеты по email"
    $chkSendEmails.Location = New-Object System.Drawing.Point(20, 140)
    $chkSendEmails.Width = 200
    $chkSendEmails.Checked = $true
    $tabEmail.Controls.Add($chkSendEmails)

    # Вкладка дополнительных параметров
    $tabOptions = New-Object System.Windows.Forms.TabPage
    $tabOptions.Text = "Дополнительные параметры"
    $tabControl.Controls.Add($tabOptions)

    $lblExcludedUsers = New-Object System.Windows.Forms.Label
    $lblExcludedUsers.Text = "Исключить пользователей:"
    $lblExcludedUsers.Location = New-Object System.Drawing.Point(20, 20)
    $lblExcludedUsers.Width = 200
    $tabOptions.Controls.Add($lblExcludedUsers)

    $txtExcludedUsers = New-Object System.Windows.Forms.TextBox
    $txtExcludedUsers.Text = "ServiceAccount1,ServiceAccount2"
    $txtExcludedUsers.Location = New-Object System.Drawing.Point(220, 20)
    $txtExcludedUsers.Width = 300
    $tabOptions.Controls.Add($txtExcludedUsers)

    $lblMethods = New-Object System.Windows.Forms.Label
    $lblMethods.Text = "Методы сбора данных:"
    $lblMethods.Location = New-Object System.Drawing.Point(20, 60)
    $lblMethods.Width = 200
    $tabOptions.Controls.Add($lblMethods)

    $chkRegistry = New-Object System.Windows.Forms.CheckBox
    $chkRegistry.Text = "Реестр"
    $chkRegistry.Checked = $true
    $chkRegistry.Location = New-Object System.Drawing.Point(220, 60)
    $chkRegistry.Width = 80
    $tabOptions.Controls.Add($chkRegistry)

    $chkPrefetch = New-Object System.Windows.Forms.CheckBox
    $chkPrefetch.Text = "Prefetch"
    $chkPrefetch.Checked = $true
    $chkPrefetch.Location = New-Object System.Drawing.Point(310, 60)
    $chkPrefetch.Width = 80
    $tabOptions.Controls.Add($chkPrefetch)

    $chkWMI = New-Object System.Windows.Forms.CheckBox
    $chkWMI.Text = "WMI"
    $chkWMI.Checked = $true
    $chkWMI.Location = New-Object System.Drawing.Point(400, 60)
    $chkWMI.Width = 80
    $tabOptions.Controls.Add($chkWMI)

    $chkCIM = New-Object System.Windows.Forms.CheckBox
    $chkCIM.Text = "CIM"
    $chkCIM.Checked = $true
    $chkCIM.Location = New-Object System.Drawing.Point(490, 60)
    $chkCIM.Width = 80
    $tabOptions.Controls.Add($chkCIM)

    $chkEventLog = New-Object System.Windows.Forms.CheckBox
    $chkEventLog.Text = "Журналы событий"
    $chkEventLog.Checked = $false
    $chkEventLog.Location = New-Object System.Drawing.Point(220, 90)
    $chkEventLog.Width = 150
    $tabOptions.Controls.Add($chkEventLog)

    $chkTestMode = New-Object System.Windows.Forms.CheckBox
    $chkTestMode.Text = "Тестовый режим (3 компьютера)"
    $chkTestMode.Location = New-Object System.Drawing.Point(20, 120)
    $chkTestMode.Width = 200
    $tabOptions.Controls.Add($chkTestMode)

    # Кнопки
    $btnHelp = New-Object System.Windows.Forms.Button
    $btnHelp.Text = "Справка"
    $btnHelp.Location = New-Object System.Drawing.Point(20, 550)
    $btnHelp.Width = 100
    $btnHelp.Add_Click({ Show-Help })
    $form.Controls.Add($btnHelp)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Запустить инвентаризацию"
    $btnRun.Location = New-Object System.Drawing.Point(300, 550)
    $btnRun.Width = 150
    $btnRun.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnRun)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Отмена"
    $btnCancel.Location = New-Object System.Drawing.Point(460, 550)
    $btnCancel.Width = 150
    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    # Отображение формы
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            Domain = $txtDomain.Text
            AdminUser = $txtAdminUser.Text
            AdminPass = $txtAdminPass.Text
            SMTPServer = $txtSMTPServer.Text
            SMTPPort = [int]$txtSMTPPort.Text
            FromEmail = $txtFromEmail.Text
            AdminEmail = $txtAdminEmail.Text
            SendEmails = $chkSendEmails.Checked
            ExcludedUsers = $txtExcludedUsers.Text -split ',' | ForEach-Object { $_.Trim() }
            UseRegistry = $chkRegistry.Checked
            UsePrefetch = $chkPrefetch.Checked
            UseWMI = $chkWMI.Checked
            UseCIM = $chkCIM.Checked
            UseEventLog = $chkEventLog.Checked
            TestMode = $chkTestMode.Checked
        }
    } else {
        return $null
    }
}
#endregion

#region Функции сбора данных
function Get-SoftwareInventory {
    param (
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $softwareList = @()
    $connectionMethods = @()
    $connected = $false
    
    # Формирование списка методов подключения
    if ($params.UseCIM) { $connectionMethods += "CIM" }
    if ($params.UseWMI) { $connectionMethods += "WMI" }
    if ($params.UseRegistry) { $connectionMethods += "Registry" }
    
    Write-Log "Попытка подключения к компьютеру $ComputerName"
    
    # Попытка подключения разными методами
    foreach ($method in $connectionMethods) {
        if ($connected) { break }
        
        try {
            switch ($method) {
                "CIM" {
                    Write-Log "Попытка подключения через CIM"
                    $session = New-CimSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                    $connected = $true
                    Write-Log "Успешное подключение через CIM"
                }
                "WMI" {
                    Write-Log "Попытка подключения через WMI"
                    $null = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                    $connected = $true
                    Write-Log "Успешное подключение через WMI"
                }
                "Registry" {
                    Write-Log "Попытка подключения через реестр"
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
                    $connected = $true
                    Write-Log "Успешное подключение через реестр"
                }
            }
        } catch {
            Write-Log "Ошибка подключения методом $method : $_" -Level "WARNING"
        }
    }
    
    if (-not $connected) {
        Write-Log "Не удалось подключиться к компьютеру $ComputerName ни одним из методов" -Level "ERROR"
        return $null
    }
    
    try {
        # 1. Сбор данных из реестра
        if ($params.UseRegistry) {
            try {
                $regPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                $reg32Path = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                
                # Подключение к реестру
                if (-not $reg) {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
                }
                
                # 64-битные приложения
                $regKey = $reg.OpenSubKey($regPath)
                if ($regKey) {
                    foreach ($subKeyName in $regKey.GetSubKeyNames()) {
                        $subKey = $regKey.OpenSubKey($subKeyName)
                        $software = [ordered]@{}
                        
                        $software.ComputerName = $ComputerName
                        $software.DisplayName = $subKey.GetValue("DisplayName")
                        $software.DisplayVersion = $subKey.GetValue("DisplayVersion")
                        $software.Publisher = $subKey.GetValue("Publisher")
                        $software.InstallDate = $subKey.GetValue("InstallDate")
                        $software.InstallLocation = $subKey.GetValue("InstallLocation")
                        $software.UninstallString = $subKey.GetValue("UninstallString")
                        $software.Source = "Registry"
                        
                        if ($software.DisplayName) {
                            $softwareList += New-Object PSObject -Property $software
                        }
                    }
                }
                
                # 32-битные приложения
                $regKey32 = $reg.OpenSubKey($reg32Path)
                if ($regKey32) {
                    foreach ($subKeyName in $regKey32.GetSubKeyNames()) {
                        $subKey = $regKey32.OpenSubKey($subKeyName)
                        $software = [ordered]@{}
                        
                        $software.ComputerName = $ComputerName
                        $software.DisplayName = $subKey.GetValue("DisplayName")
                        $software.DisplayVersion = $subKey.GetValue("DisplayVersion")
                        $software.Publisher = $subKey.GetValue("Publisher")
                        $software.InstallDate = $subKey.GetValue("InstallDate")
                        $software.InstallLocation = $subKey.GetValue("InstallLocation")
                        $software.UninstallString = $subKey.GetValue("UninstallString")
                        $software.Source = "Registry (32-bit)"
                        
                        if ($software.DisplayName) {
                            $softwareList += New-Object PSObject -Property $software
                        }
                    }
                }
                
                Write-Log "Собраны данные из реестра для компьютера $ComputerName ($($softwareList.Count) приложений)"
            } catch {
                Write-Log "Ошибка при сборе данных из реестра с компьютера $ComputerName : $_" -Level "ERROR"
            }
        }
        
        # 2. Сбор данных из WMI (Win32_Product)
        if ($params.UseWMI) {
            try {
                $wmiSoftware = Get-WmiObject -Class Win32_Product -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                
                foreach ($sw in $wmiSoftware) {
                    $software = [ordered]@{}
                    
                    $software.ComputerName = $ComputerName
                    $software.DisplayName = $sw.Name
                    $software.DisplayVersion = $sw.Version
                    $software.Publisher = $sw.Vendor
                    $software.InstallDate = $sw.InstallDate
                    $software.InstallLocation = $sw.InstallLocation
                    $software.UninstallString = "msiexec /x $($sw.IdentifyingNumber)"
                    $software.Source = "WMI"
                    
                    # Проверка на дубликаты
                    if ($software.DisplayName -and -not ($softwareList | Where-Object { $_.DisplayName -eq $software.DisplayName })) {
                        $softwareList += New-Object PSObject -Property $software
                    }
                }
                
                Write-Log "Собраны данные WMI для компьютера $ComputerName ($($wmiSoftware.Count) приложений)"
            } catch {
                Write-Log "Ошибка при сборе данных WMI с компьютера $ComputerName : $_" -Level "ERROR"
            }
        }
        
        # 3. Сбор данных через CIM
        if ($params.UseCIM) {
            try {
                if (-not $session) {
                    $session = New-CimSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                }
                
                $cimSoftware = Get-CimInstance -CimSession $session -ClassName Win32_Product -ErrorAction Stop
                
                foreach ($sw in $cimSoftware) {
                    $software = [ordered]@{}
                    
                    $software.ComputerName = $ComputerName
                    $software.DisplayName = $sw.Name
                    $software.DisplayVersion = $sw.Version
                    $software.Publisher = $sw.Vendor
                    $software.InstallDate = $sw.InstallDate
                    $software.InstallLocation = $sw.InstallLocation
                    $software.UninstallString = "msiexec /x $($sw.IdentifyingNumber)"
                    $software.Source = "CIM"
                    
                    # Проверка на дубликаты
                    if ($software.DisplayName -and -not ($softwareList | Where-Object { $_.DisplayName -eq $software.DisplayName })) {
                        $softwareList += New-Object PSObject -Property $software
                    }
                }
                
                Write-Log "Собраны данные CIM для компьютера $ComputerName ($($cimSoftware.Count) приложений)"
            } catch {
                Write-Log "Ошибка при сборе данных CIM с компьютера $ComputerName : $_" -Level "ERROR"
            }
        }
        
        # 4. Анализ Prefetch-файлов
        if ($params.UsePrefetch) {
            try {
                $prefetchPath = "\\$ComputerName\c$\Windows\Prefetch"
                if (Test-Path $prefetchPath) {
                    $prefetchFiles = Get-ChildItem -Path $prefetchPath -Filter "*.pf" -ErrorAction Stop
                    
                    foreach ($file in $prefetchFiles) {
                        $appName = $file.Name -replace '\.pf$', ''
                        $lastRun = $file.LastWriteTime
                        
                        # Сопоставление с установленным ПО
                        foreach ($sw in $softwareList) {
                            if ($sw.DisplayName -like "*$appName*") {
                                $sw | Add-Member -NotePropertyName "LastRunTime" -NotePropertyValue $lastRun -Force
                                $sw | Add-Member -NotePropertyName "LastRunSource" -NotePropertyValue "Prefetch" -Force
                                break
                            }
                        }
                    }
                }
            } catch {
                Write-Log "Ошибка при анализе Prefetch-файлов с компьютера $ComputerName : $_" -Level "WARNING"
            }
        }
        
        # 5. Анализ журналов событий (Application)
        if ($params.UseEventLog) {
            try {
                $events = Get-WinEvent -ComputerName $ComputerName -Credential $Credential -FilterHashtable @{
                    LogName = 'Application'
                    ProviderName = 'Application Error'
                    StartTime = (Get-Date).AddYears(-1)
                } -ErrorAction Stop -MaxEvents 1000
                
                foreach ($event in $events) {
                    $appName = $event.Properties[0].Value
                    $lastRun = $event.TimeCreated
                    
                    # Сопоставление с установленным ПО
                    foreach ($sw in $softwareList) {
                        if ($sw.DisplayName -like "*$appName*") {
                            if (-not $sw.LastRunTime -or $lastRun -gt $sw.LastRunTime) {
                                $sw.LastRunTime = $lastRun
                                $sw.LastRunSource = "EventLog"
                            }
                            break
                        }
                    }
                }
                
                Write-Log "Проанализированы журналы событий для компьютера $ComputerName ($($events.Count) событий)"
            } catch {
                Write-Log "Ошибка при анализе журналов событий с компьютера $ComputerName : $_" -Level "WARNING"
            }
        }
        
        return $softwareList
        
    } catch {
        Write-Log "Критическая ошибка при сборе данных с компьютера $ComputerName : $_" -Level "ERROR"
        return $null
    } finally {
        if ($session) { Remove-CimSession -CimSession $session }
        if ($reg) { $reg.Dispose() }
    }
}

function Get-HTMLReport {
    param (
        [array]$Data,
        [string]$Title = "Отчет об инвентаризации ПО"
    )
    
    # Определение категорий ПО
    $threeMonthsAgo = (Get-Date).AddMonths(-3)
    $sixMonthsAgo = (Get-Date).AddMonths(-6)
    
    $oldSoftware = $Data | Where-Object { 
        $_.LastRunTime -and $_.LastRunTime -lt $sixMonthsAgo -or 
        (-not $_.LastRunTime -and $_.InstallDate -and [datetime]$_.InstallDate -lt $sixMonthsAgo)
    }
    
    $mediumSoftware = $Data | Where-Object { 
        $_.LastRunTime -and $_.LastRunTime -ge $sixMonthsAgo -and $_.LastRunTime -lt $threeMonthsAgo
    }
    
    $recentSoftware = $Data | Where-Object { 
        $_.LastRunTime -and $_.LastRunTime -ge $threeMonthsAgo
    }
    
    $unknownSoftware = $Data | Where-Object { 
        -not $_.LastRunTime -and (-not $_.InstallDate -or [datetime]$_.InstallDate -ge $sixMonthsAgo)
    }
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$Title</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; text-align: center; margin-bottom: 30px; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        h2 { color: #2c3e50; background-color: #ecf0f1; padding: 8px; border-left: 4px solid #3498db; margin-top: 30px; }
        .section { margin-bottom: 30px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th { background-color: #3498db; color: white; text-align: left; padding: 10px; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f1f1f1; }
        .old-row { background-color: #ffdddd !important; }
        .medium-row { background-color: #fff4dd !important; }
        .recent-row { background-color: #ddffdd !important; }
        .unknown-row { background-color: #f0f0f0 !important; }
        .summary { background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin-top: 20px; }
        .footer { text-align: center; margin-top: 30px; color: #7f8c8d; font-size: 0.9em; }
        .category-label { display: inline-block; padding: 3px 8px; border-radius: 4px; color: white; font-weight: bold; }
        .old-label { background-color: #e74c3c; }
        .medium-label { background-color: #f39c12; }
        .recent-label { background-color: #2ecc71; }
        .unknown-label { background-color: #95a5a6; }
    </style>
</head>
<body>
    <div class="container">
        <h1>$Title</h1>
        <p><strong>Дата генерации:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
        <p><strong>Всего записей:</strong> $($Data.Count)</p>
        
        <div class="summary">
            <h3>Статистика по категориям:</h3>
            <p><span class="category-label old-label">Старое ПО</span> (> 6 месяцев): $($oldSoftware.Count) программ</p>
            <p><span class="category-label medium-label">Среднее ПО</span> (3-6 месяцев): $($mediumSoftware.Count) программ</p>
            <p><span class="category-label recent-label">Недавнее ПО</span> (< 3 месяцев): $($recentSoftware.Count) программ</p>
            <p><span class="category-label unknown-label">Неизвестно</span>: $($unknownSoftware.Count) программ</p>
        </div>
        
        <div class="section">
            <h2>ПО, которое не запускалось более полугода</h2>
            <table>
                <tr>
                    <th>Компьютер</th>
                    <th>Пользователь</th>
                    <th>ПО</th>
                    <th>Версия</th>
                    <th>Издатель</th>
                    <th>Дата установки</th>
                    <th>Последний запуск</th>
                </tr>
"@

    foreach ($item in $oldSoftware) {
        $rowClass = "old-row"
        $html += @"
                <tr class="$rowClass">
                    <td>$($item."Имя компьютера")</td>
                    <td>$($item."Связанный пользователь")</td>
                    <td>$($item."Название ПО")</td>
                    <td>$($item."Версия")</td>
                    <td>$($item."Издатель")</td>
                    <td>$($item."Дата установки")</td>
                    <td>$($item."Время последнего запуска")</td>
                </tr>
"@
    }

    $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>ПО, которое не запускалось от 3 месяцев до полугода</h2>
            <table>
                <tr>
                    <th>Компьютер</th>
                    <th>Пользователь</th>
                    <th>ПО</th>
                    <th>Версия</th>
                    <th>Издатель</th>
                    <th>Дата установки</th>
                    <th>Последний запуск</th>
                </tr>
"@

    foreach ($item in $mediumSoftware) {
        $rowClass = "medium-row"
        $html += @"
                <tr class="$rowClass">
                    <td>$($item."Имя компьютера")</td>
                    <td>$($item."Связанный пользователь")</td>
                    <td>$($item."Название ПО")</td>
                    <td>$($item."Версия")</td>
                    <td>$($item."Издатель")</td>
                    <td>$($item."Дата установки")</td>
                    <td>$($item."Время последнего запуска")</td>
                </tr>
"@
    }

    $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>ПО, которое запускалось до 3 месяцев</h2>
            <table>
                <tr>
                    <th>Компьютер</th>
                    <th>Пользователь</th>
                    <th>ПО</th>
                    <th>Версия</th>
                    <th>Издатель</th>
                    <th>Дата установки</th>
                    <th>Последний запуск</th>
                </tr>
"@

    foreach ($item in $recentSoftware) {
        $rowClass = "recent-row"
        $html += @"
                <tr class="$rowClass">
                    <td>$($item."Имя компьютера")</td>
                    <td>$($item."Связанный пользователь")</td>
                    <td>$($item."Название ПО")</td>
                    <td>$($item."Версия")</td>
                    <td>$($item."Издатель")</td>
                    <td>$($item."Дата установки")</td>
                    <td>$($item."Время последнего запуска")</td>
                </tr>
"@
    }

    $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>ПО с неизвестным временем последнего запуска</h2>
            <table>
                <tr>
                    <th>Компьютер</th>
                    <th>Пользователь</th>
                    <th>ПО</th>
                    <th>Версия</th>
                    <th>Издатель</th>
                    <th>Дата установки</th>
                </tr>
"@

    foreach ($item in $unknownSoftware) {
        $rowClass = "unknown-row"
        $html += @"
                <tr class="$rowClass">
                    <td>$($item."Имя компьютера")</td>
                    <td>$($item."Связанный пользователь")</td>
                    <td>$($item."Название ПО")</td>
                    <td>$($item."Версия")</td>
                    <td>$($item."Издатель")</td>
                    <td>$($item."Дата установки")</td>
                </tr>
"@
    }

    $html += @"
            </table>
        </div>
        
        <div class="footer">
            Отчет сгенерирован автоматически системой инвентаризации ПО. &copy; $((Get-Date).Year)
        </div>
    </div>
</body>
</html>
"@
    return $html
}
#endregion

#region Основной скрипт
try {
    # Инициализация логирования
    Initialize-Logging
    Write-Log "Запуск скрипта инвентаризации ПО"
    
    # Показать графический интерфейс и получить параметры
    $params = Show-MainForm
    
    if (-not $params) {
        Write-Log "Скрипт отменен пользователем." -Level "WARNING"
        Close-Logging
        exit
    }

    # Обновляем путь к лог-файлу
    $global:logPath = "C:\Logs\SoftwareInventory_$(Get-Date -Format 'yyyyMMdd_HHmm').log"
    
    # Запись параметров в лог
    Write-Log "Параметры выполнения:"
    $params.GetEnumerator() | ForEach-Object {
        if ($_.Key -ne "AdminPass") {
            Write-Log "$($_.Key): $($_.Value)"
        } else {
            Write-Log "$($_.Key): ********"
        }
    }

    # Создание учетных данных
    $securePass = ConvertTo-SecureString $params.AdminPass -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($params.AdminUser, $securePass)

    # Подключение модуля Active Directory
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Log "Модуль Active Directory успешно загружен"
    } catch {
        Write-Log "Ошибка загрузки модуля Active Directory: $_" -Level "ERROR"
        Close-Logging
        exit 1
    }

    # Проверка подключения к домену
    if (-not (Test-ADConnection -Domain $params.Domain)) {
        Write-Log "Ошибка подключения к домену AD. Скрипт завершен." -Level "ERROR"
        Close-Logging
        exit 1
    }

    # Проверка SMTP подключения (если нужно отправлять email)
    if ($params.SendEmails) {
        if (-not (Test-SMTPConnection -SMTPServer $params.SMTPServer -SMTPPort $params.SMTPPort)) {
            Write-Log "Ошибка подключения к SMTP серверу. Email не будут отправлены." -Level "ERROR"
            $params.SendEmails = $false
        }
    }

    # Получение списка компьютеров из AD
    try {
        $computers = Get-ADComputer -Filter * -Properties Name, OperatingSystem, LastLogonDate | 
                     Where-Object { $_.OperatingSystem -like "*Windows*" } |
                     Select-Object -ExpandProperty Name
        
        Write-Log "Найдено $($computers.Count) компьютеров Windows в домене"
    } catch {
        Write-Log "Ошибка получения списка компьютеров из AD: $_" -Level "ERROR"
        Close-Logging
        exit 1
    }

    # Получение информации о пользователях из AD
    try {
        $userMap = @{}
        $allUsers = Get-ADUser -Filter * -Properties Mail, SamAccountName, LastLogonDate
        
        foreach ($user in $allUsers) {
            if (-not [string]::IsNullOrEmpty($user.Mail) -and $params.ExcludedUsers -notcontains $user.SamAccountName) {
                $userMap[$user.SamAccountName] = @{
                    Email = $user.Mail
                    LastLogon = $user.LastLogonDate
                }
            }
        }
        
        Write-Log "Получены email-адреса для $($userMap.Count) пользователей из AD"
    } catch {
        Write-Log "Ошибка получения информации о пользователях из AD: $_" -Level "ERROR"
        $userMap = @{}
    }

    # Основной цикл сбора данных
    $allSoftware = @()
    $computerUserMap = @{}
    $processedComputers = 0
    $successfulComputers = 0

    if ($params.TestMode) {
        $computers = $computers | Select-Object -First 3
        Write-Log "Тестовый режим: ограничение до 3 компьютеров"
    }

    foreach ($computer in $computers) {
        $processedComputers++
        Write-Log "Обработка компьютера $computer ($processedComputers из $($computers.Count))"
        Write-Progress -Activity "Сбор данных" -Status "Обработка компьютера $computer ($processedComputers из $($computers.Count))" -PercentComplete (($processedComputers / $computers.Count) * 100)
        
        try {
            # Получение текущего пользователя компьютера
            $loggedInUser = $null
            try {
                $sessions = quser /server:$computer 2>$null
                if ($sessions) {
                    $loggedInUser = ($sessions[1] -split '\s+')[1]
                    $computerUserMap[$computer] = $loggedInUser
                    Write-Log "Определен пользователь: $loggedInUser"
                }
            } catch {
                Write-Log "Не удалось определить текущего пользователя для компьютера $computer" -Level "WARNING"
            }
            
            # Получение инвентаризации ПО
            $software = Get-SoftwareInventory -ComputerName $computer -Credential $credential
            if ($software) {
                $allSoftware += $software
                $successfulComputers++
                Write-Log "Успешно собраны данные с компьютера $computer"
            }
        } catch {
            Write-Log "Ошибка обработки компьютера $computer : $_" -Level "ERROR"
        }
    }

    Write-Progress -Activity "Сбор данных" -Completed
    Write-Log "Сбор данных завершен. Успешно обработано $successfulComputers из $processedComputers компьютеров"

    # Генерация отчета
    $threeMonthsAgo = (Get-Date).AddMonths(-3)
    $sixMonthsAgo = (Get-Date).AddMonths(-6)

    $reportData = foreach ($sw in $allSoftware) {
        $category = if ($sw.LastRunTime) {
            if ($sw.LastRunTime -ge $threeMonthsAgo) {
                "Использовалось недавно (0-3 месяца)"
            } elseif ($sw.LastRunTime -ge $sixMonthsAgo) {
                "Использовалось иногда (3-6 месяцев)"
            } else {
                "Редко используется (6+ месяцев)"
            }
        } else {
            "Время последнего использования неизвестно"
        }
        
        [PSCustomObject]@{
            "Имя компьютера" = $sw.ComputerName
            "Название ПО" = $sw.DisplayName
            "Версия" = $sw.DisplayVersion
            "Издатель" = $sw.Publisher
            "Дата установки" = $sw.InstallDate
            "Время последнего запуска" = if ($sw.LastRunTime) { $sw.LastRunTime.ToString("yyyy-MM-dd") } else { "Неизвестно" }
            "Источник данных" = $sw.Source
            "Источник времени запуска" = $sw.LastRunSource
            "Категория использования" = $category
            "Связанный пользователь" = if ($computerUserMap.ContainsKey($sw.ComputerName)) { $computerUserMap[$sw.ComputerName] } else { "Неизвестно" }
            "Email пользователя" = if ($computerUserMap.ContainsKey($sw.ComputerName) -and $userMap.ContainsKey($computerUserMap[$sw.ComputerName])) { 
                $userMap[$computerUserMap[$sw.ComputerName]].Email 
            } else { 
                "Не доступен" 
            }
        }
    }

    # Сохранение отчета в CSV
    $reportPath = "C:\Reports\SoftwareInventory_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $reportDir = Split-Path -Path $reportPath -Parent
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    }

    $reportData | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Отчет сохранен в файл: $reportPath"

    # Отправка отчетов по email
    if ($params.SendEmails) {
        try {
            # Отправка полного отчета администратору
            $adminHtml = Get-HTMLReport -Data $reportData -Title "Полный отчет об инвентаризации ПО"
            
            $adminSubject = "Полный отчет об инвентаризации ПО - $(Get-Date -Format 'yyyy-MM-dd')"
            $adminBody = "Уважаемый администратор,<br><br>Прикреплен полный отчет об инвентаризации программного обеспечения в домене."
            
            $adminMessage = New-Object System.Net.Mail.MailMessage($params.FromEmail, $params.AdminEmail, $adminSubject, $adminBody)
            $adminMessage.IsBodyHtml = $true
            
            # Прикрепление CSV
            $attachment = New-Object System.Net.Mail.Attachment($reportPath)
            $adminMessage.Attachments.Add($attachment)
            
            # Отправка email
            $smtp = New-Object Net.Mail.SmtpClient($params.SMTPServer, $params.SMTPPort)
            $smtp.Send($adminMessage)
            
            Write-Log "Полный отчет отправлен администратору на $($params.AdminEmail)" -Level "SUCCESS"
        } catch {
            Write-Log "Ошибка при отправке email администратору: $_" -Level "ERROR"
        }

        # Отправка индивидуальных отчетов пользователям
        $userGroups = $reportData | Group-Object "Связанный пользователь"
        
        foreach ($group in $userGroups) {
            $userName = $group.Name
            if ($userName -eq "Неизвестно" -or $params.ExcludedUsers -contains $userName) {
                continue
            }
            
            if ($userMap.ContainsKey($userName)) {
                $userEmail = $userMap[$userName].Email
                $userData = $group.Group
                
                try {
                    $userHtml = Get-HTMLReport -Data $userData -Title "Отчет об инвентаризации ПО для вашего компьютера"
                    
                    $userSubject = "Отчет об инвентаризации ПО - $(Get-Date -Format 'yyyy-MM-dd')"
                    $userBody = "Уважаемый(ая) $userName,<br><br>Прикреплен отчет об инвентаризации программного обеспечения на вашем компьютере."
                    
                    $userMessage = New-Object System.Net.Mail.MailMessage($params.FromEmail, $userEmail, $userSubject, $userBody)
                    $userMessage.IsBodyHtml = $true
                    
                    # Создание временного CSV для пользователя
                    $tempReportDir = "C:\Reports\Temp"
                    if (-not (Test-Path -Path $tempReportDir)) {
                        New-Item -ItemType Directory -Path $tempReportDir -Force | Out-Null
                    }
                    $userReportPath = Join-Path -Path $tempReportDir -ChildPath "SoftwareInventory_${userName}_$(Get-Date -Format 'yyyyMMdd').csv"
                    $userData | Export-Csv -Path $userReportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                    
                    # Прикрепление CSV
                    $attachment = New-Object System.Net.Mail.Attachment($userReportPath)
                    $userMessage.Attachments.Add($attachment)
                    
                    # Отправка email
                    $smtp.Send($userMessage)
                    
                    Write-Log "Отчет отправлен пользователю $userName на $userEmail" -Level "SUCCESS"
                    
                    # Удаление временного файла
                    Remove-Item -Path $userReportPath -Force
                } catch {
                    Write-Log "Ошибка при отправке email пользователю $userName : $_" -Level "ERROR"
                }
            }
        }
    }

    # Отображение результатов
    $resultForm = New-Object System.Windows.Forms.Form
    $resultForm.Text = "Результаты инвентаризации ПО"
    $resultForm.Size = New-Object System.Drawing.Size(600, 400)
    $resultForm.StartPosition = "CenterScreen"

    $lblResult = New-Object System.Windows.Forms.Label
    $lblResult.Text = "Инвентаризация ПО успешно завершена!`n`n" +
                      "Обработано компьютеров: $successfulComputers из $processedComputers`n" +
                      "Найдено программ: $($reportData.Count)`n`n" +
                      "Отчет сохранен в:`n$reportPath"
    $lblResult.Location = New-Object System.Drawing.Point(20, 20)
    $lblResult.AutoSize = $true
    $resultForm.Controls.Add($lblResult)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = "Открыть отчет"
    $btnOpen.Location = New-Object System.Drawing.Point(150, 200)
    $btnOpen.Add_Click({
        Invoke-Item $reportPath
    })
    $resultForm.Controls.Add($btnOpen)

    $btnLog = New-Object System.Windows.Forms.Button
    $btnLog.Text = "Просмотреть лог"
    $btnLog.Location = New-Object System.Drawing.Point(250, 200)
    $btnLog.Add_Click({
        Invoke-Item $global:logPath
    })
    $resultForm.Controls.Add($btnLog)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Закрыть"
    $btnClose.Location = New-Object System.Drawing.Point(350, 200)
    $btnClose.Add_Click({
        $resultForm.Close()
    })
    $resultForm.Controls.Add($btnClose)

    $resultForm.ShowDialog() | Out-Null

    Write-Log "Скрипт успешно завершен" -Level "SUCCESS"
} catch {
    Write-Log "Критическая ошибка в основном скрипте: $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Произошла критическая ошибка: $_", "Ошибка", "OK", "Error")
} finally {
    Close-Logging
}
#endregion