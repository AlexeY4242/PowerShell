# Загрузка Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Главный экран
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Запуск и редактирование скриптов'
$form.Size = New-Object System.Drawing.Size(1500,900)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::LightSteelBlue

# Заголовок выбора файла
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Выберите файл или отредактируйте скрипт:'
$label.Size = New-Object System.Drawing.Size(680,20)
$label.Location = New-Object System.Drawing.Point(20,20)
$label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$label.ForeColor = [System.Drawing.Color]::DarkSlateGray

# Заголовок редактирование кода
$labelRight = New-Object System.Windows.Forms.Label
$labelRight.Text = 'Редактирование кода PowerShell и Python:'
$labelRight.Size = New-Object System.Drawing.Size(680,20)
$labelRight.Location = New-Object System.Drawing.Point(730,20)
$labelRight.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$labelRight.ForeColor = [System.Drawing.Color]::DarkSlateGray

# Список скриптов
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Size = New-Object System.Drawing.Size(680,640)
$listBox.Location = New-Object System.Drawing.Point(20,50)
$listBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$listBox.HorizontalScrollbar = $true
$listBox.BorderStyle = 'FixedSingle'

# Путь
$scriptPath = "C:\Users\ALVKonstantinov\Desktop\Задания на год"
Get-ChildItem -Path $scriptPath -Recurse | Where-Object {
    $_.Extension -eq '.ps1' -or $_.Extension -eq '.py'
} | ForEach-Object {
    $listBox.Items.Add($_.FullName)
}

# Поле для редактирования скрипта
$richTextBox = New-Object System.Windows.Forms.RichTextBox
$richTextBox.Size = New-Object System.Drawing.Size(680,700)
$richTextBox.Location = New-Object System.Drawing.Point(730,50)
$richTextBox.Font = New-Object System.Drawing.Font("Consolas", 12)
$richTextBox.BorderStyle = 'FixedSingle'
$richTextBox.BackColor = [System.Drawing.Color]::LightGoldenrodYellow

#Редактирование информации в поле
$listBox.Add_SelectedIndexChanged({
    $selectedFile = $listBox.SelectedItem
    if ($selectedFile -and (Test-Path -Path $selectedFile)) {
        try {
            $content = Get-Content -Path $selectedFile -Raw -Encoding UTF8    # Чтение файла с кодировкой UTF-8
            $richTextBox.Text = $content                                       # Запись текста в RichTextBox
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Ошибка при открытии файла '" + $selectedFile + "'.", "Ошибка", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Подсветка ключевых слов и операторов
$richTextBox.Add_TextChanged({
    # Сохраняем текущее положение курсора
    $cursorPosition = $richTextBox.SelectionStart
    
    # Списки для подсветки
    $keywords = @('if', 'else', 'try', 'catch', 'foreach', 'function', 'return', 'Write-Host', 'switch', 'break')
    $operators = @('-eq', '-ne', '-lt', '-gt', '-and', '-or', '=', '+', '-', '*', '/')

    # Задаем цвета для подсветки
    $commentColor = [System.Drawing.Color]::Green
    $keywordColor = [System.Drawing.Color]::Blue
    $operatorColor = [System.Drawing.Color]::Purple
    $defaultColor = [System.Drawing.Color]::Black

    # Очищаем форматирование
    $richTextBox.SelectAll()
    $richTextBox.SelectionColor = $defaultColor
    $richTextBox.DeselectAll()
    
    # Подсвечиваем ключевые слова
    foreach ($keyword in $keywords) {
        $matches = [regex]::Matches($richTextBox.Text, "\b$keyword\b")
        foreach ($match in $matches) {
            $richTextBox.Select($match.Index, $match.Length)
            $richTextBox.SelectionColor = $keywordColor
        }
    }

    # Подсвечиваем операторы
    foreach ($operator in $operators) {
        $matches = [regex]::Matches($richTextBox.Text, [regex]::Escape($operator))
        foreach ($match in $matches) {
            $richTextBox.Select($match.Index, $match.Length)
            $richTextBox.SelectionColor = $operatorColor
        }
    }

    # Подсвечиваем комментарии
    $comments = [regex]::Matches($richTextBox.Text, "#.*")
    foreach ($comment in $comments) {
        $richTextBox.Select($comment.Index, $comment.Length)
        $richTextBox.SelectionColor = $commentColor
    }

    # Восстанавливаем положение курсора
    $richTextBox.SelectionStart = $cursorPosition
    $richTextBox.SelectionLength = 0
    $richTextBox.ScrollToCaret()
})


# Кнопка Выполнить
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = 'Выполнить'
$buttonRun.Size = New-Object System.Drawing.Size(140,50)
$buttonRun.Location = New-Object System.Drawing.Point(350,700)
$buttonRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonRun.BackColor = [System.Drawing.Color]::ForestGreen
$buttonRun.ForeColor = [System.Drawing.Color]::White
$buttonRun.FlatStyle = 'Flat'
$buttonRun.Cursor = [System.Windows.Forms.Cursors]::Hand

#Выполнить
$buttonRun.Add_Click({
    $script = $listBox.SelectedItem
    if ($script) {
        $fullScriptPath = $script  # Поскольку мы уже добавили полный путь в ListBox
        if (Test-Path -Path $fullScriptPath) {
            Write-Host "Запуск скрипта $fullScriptPath"
            # Запуск PowerShell-скрипта
            Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$fullScriptPath`""
        } else {
            [System.Windows.Forms.MessageBox]::Show('Файл не найден.', 'Ошибка', 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show('Выберите файл для выполнения.', 'Ошибка', 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Кнопка Редактирование
$buttonEdit = New-Object System.Windows.Forms.Button
$buttonEdit.Text = 'Редактировать файл'
$buttonEdit.Size = New-Object System.Drawing.Size(140,50)
$buttonEdit.Location = New-Object System.Drawing.Point(150,700)
$buttonEdit.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonEdit.BackColor = [System.Drawing.Color]::DodgerBlue
$buttonEdit.ForeColor = [System.Drawing.Color]::White
$buttonEdit.FlatStyle = 'Flat'
$buttonEdit.Cursor = [System.Windows.Forms.Cursors]::Hand

 # Редактирование файлов
$buttonEdit.Add_Click({
    $script = $listBox.SelectedItem
    if ($script) {
        $fullScriptPath = $script  # Поскольку мы уже добавили полный путь в ListBox

        # Определение программы для редактирования в зависимости от расширения
        switch ([System.IO.Path]::GetExtension($fullScriptPath)) {
            ".ps1" { $editor = "PowerShell_ISE.exe" }  # PowerShell для .ps1
            ".py" { $editor = "C:\Program Files\Microsoft VS Code\Code.exe" }  # Sublime Text для .py (проверьте путь)
            default { 
                [System.Windows.Forms.MessageBox]::Show('Не поддерживаемое расширение файла.', 'Ошибка', 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        

        # Проверим существования файла
      if (Test-Path -Path $fullScriptPath) {
           Write-Host "Открываем файл $fullScriptPath в $editor"
            # Экранируем полный путь к файлу
           Start-Process -FilePath $editor -ArgumentList "`"$fullScriptPath`""
       } else {
           [System.Windows.Forms.MessageBox]::Show('Файл не найден.', 'Ошибка', 
           [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
       }
   } else {
       [System.Windows.Forms.MessageBox]::Show('Пожалуйста, выберите файл для редактирования.', 'Ошибка', 
       [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
   }
})
# Кнопка Сохранить
$buttonSave = New-Object System.Windows.Forms.Button
$buttonSave.Text = 'Сохранить изменения'
$buttonSave.Size = New-Object System.Drawing.Size(140,40)
$buttonSave.Location = New-Object System.Drawing.Point(1000,767)  # Расположим кнопку под полем редактирования
$buttonSave.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonSave.BackColor = [System.Drawing.Color]::DarkOrange  # Изменим цвет кнопки
$buttonSave.ForeColor = [System.Drawing.Color]::White
$buttonSave.FlatStyle = 'Flat'
$buttonSave.Cursor = [System.Windows.Forms.Cursors]::Hand

#Сохранение файлов
$buttonSave.Add_Click({
    $selectedFile = $listBox.SelectedItem
    if ($selectedFile) {
        if (Test-Path -Path $selectedFile) {
            try {
                $richTextBox.Text | Set-Content -Path $selectedFile -Encoding UTF8    # Сохранение файла с кодировкой UTF-8
                [System.Windows.Forms.MessageBox]::Show('Изменения сохранены.', 'Успех', 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show('Ошибка при сохранении файла.', 'Ошибка', 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show('Файл не найден.', 'Ошибка', 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show('Выберите файл для сохранения.', 'Ошибка', 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Добавление компонентов в форму

$form.Controls.Add($label)
$form.Controls.Add($labelRight)
$form.Controls.Add($listBox)
$form.Controls.Add($buttonRun)
$form.Controls.Add($buttonEdit)
$form.Controls.Add($buttonSave)
$form.Controls.Add($richTextBox)


# Запуск формы
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
