Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function ShowErrorNExit($ErrorCode) {
    Write-Host $ErrorText' '$ErrorCode'`n' -ForegroundColor Red
    Write-Host $PaKText'`n'
    [void] $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
}

$Languages = @(@('en-US', 'es-ES', 'it-IT', 'zh-CN', 'zh-TW', 'ko-KR', 'de-DE', 'pl-PL', 'pt-BR', 'ru-RU', 'tr-TR', 'fr-FR', 'cs-CZ', 'ja-JP'), @())
$OKText = @('OK', 'Aceptar', 'OK', '确定', '確定', '확인', 'OK', 'OK', 'OK', 'OK', 'Tamam', 'OK', 'OK', 'OK')
$CancelText = @('Cancel', 'Cancelar', 'Annulla', '取消', '取消', '취소', 'Abbrechen', 'Anuluj', 'Cancelar', 'Отмена', 'İptal', 'Annuler', 'Storno', 'キャンセル')
$WindowTitle = @('Offline package settings',
                 'Configuración de paquetes sin conexión',
                 'Impostazioni del pacchetto offline',
                 '离线包设置',
                 '離線包設置',
                 '오프라인 패키지 설정',
                 'Offline-Paketeinstellungen',
                 'Ustawienia pakietu offline',
                 'Configurações de pacote off-line',
                 'Настройки автономного пакета',
                 'Çevrimdışı paket ayarları',
                 'Paramètres du package hors ligne',
                 'Nastavení offline balíku',
                 'オフラインパッケージの設定')
$EditionTitle = @('Visual Studio Edition:',
                  'Edición de Visual Studio:',
                  'Edizione di Visual Studio:',
                  'Visual Studio 版本:',
                  'Visual Studio 的版本:',
                  'Visual Studio 버전:',
                  'Visual Studio-Edition:',
                  'Wersja Visual Studio:',
                  'Edição do Visual Studio:',
                  'Редакция Visual Studio:',
                  'Visual Studio sürümü:',
                  'Édition de Visual Studio:',
                  'Edice Visual Studio:',
                  'Visual Studio エディション:')
$LangTitle = @('Needed languages:',
               'Seleccione los idiomas necesarios:',
               'Seleziona le lingue necessarie:',
               '选择需要的语言：',
               '選擇需要的語言：',
               '필요한 언어 선택 :',
               'Gewünschte sprachen auswählen:',
               'Wybierz potrzebne języki:',
               'Selecione os idiomas necessários:',
               'Выберите необходимые языки:',
               'Gerekli dilleri seçin:',
               'Sélectionnez les langues nécessaires :',
               'Vyberte potřebné jazyky:',
               '必要な言語を選択してください:')
$ErrorText = @('Error', 'Error', 'Errore', '错误', '錯誤', '오류', 'Error', 'Błąd', 'Erro', 'Ошибка', 'Hata', 'Erreur', 'Chyba', 'エラー')
$PaKText = @('Press any key...',
             'Presiona cualquier tecla...',
             'Premere un tasto qualsiasi...',
             '按任意键...',
             '按任意鍵...',
             '아무 키나 누르시오...',
             'Drücke irgendeine taste...',
             'Naciśnij dowolny klawisz...',
             'Pressione qualquer tecla...',
             'Нажмите любую клавишу...',
             'Herhangi bir tuşa basın...',
             'Appuyez sur une touche...',
             'Stiskněte libovolnou klávesu...',
             'どれかキーを押してください...')

$Cult = Get-Culture -ListAvailable
$LngName = Get-Culture
$LngIndex = $Languages[0].IndexOf($LngName.Name)
if ($LngIndex -eq -1) {
    $LngIndex = 0
}
foreach ($item in $Languages[0]) {
    $result = $Cult | Where-Object { $_.Name -like $item }
    $Languages[1] += (Get-Culture).TextInfo.ToTitleCase($result.DisplayName)
}

$Editions = @()
$AllURLs = Invoke-WebRequest "https://visualstudio.microsoft.com/${$LngName.Name}/downloads/" | Select-String -InputObject { $_.Content } -Pattern 'http\S+studio/\?sku\S+(?=\")' -AllMatches
$UniqURLs = $AllURLs.Matches.Value | Select-Object -Unique
foreach ($item in $UniqURLs) {
    $result = Select-String -InputObject $item -Pattern '(?<=sku\=)\w+'
    $Editions += $result.Matches.Value
}
<#
-------------------- Window form --------------------
#>
$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle[$LngIndex]
$form.Size = New-Object System.Drawing.Size(400,370)
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 20)
$label1.Size = New-Object System.Drawing.Size(150, 20)
$label1.Text = $EditionTitle[$LngIndex]
$form.Controls.Add($label1)

$listboxEd = New-Object System.Windows.Forms.ComboBox
$listboxEd.Location = New-Object System.Drawing.Point(170, 18)
$listboxEd.Size = New-Object System.Drawing.Size(150, 20)
foreach ($ed in $Editions) {
    [void] $listBoxEd.Items.Add($ed)
}
$listboxEd.DropDownStyle = 'DropDownList'
$listboxEd.SelectedIndex = 2
$form.Controls.Add($listboxEd)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20, 50)
$label2.Size = New-Object System.Drawing.Size(250, 20)
$label2.Text = 'Необходимые языки:'
$form.Controls.Add($label2)

$listboxLng = New-Object System.Windows.Forms.ListBox
$listboxLng.Location = New-Object System.Drawing.Point(20, 70)
$listboxLng.Size = New-Object System.Drawing.Size(220, 20)
foreach ($lng in $Languages[1]) {
    [void] $listboxLng.Items.Add($lng)
}
$listboxLng.SetSelected($LngIndex, $true);
$listboxLng.SelectionMode = 'MultiExtended'
$listboxLng.Height = 250
$form.Controls.Add($listboxLng)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(260, 240)
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Text = $OKText[$LngIndex]
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(260, 280)
$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
$cancelButton.Text = $CancelText[$LngIndex]
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
$form.Topmost = $true

$result = $form.ShowDialog()
if ($result -ne $OKText[$LngIndex]) {
    exit
}
<#
-------------------- Main code --------------------
#>
$Edition = $Editions[$listboxEd.SelectedIndex]
$Exe = ".\$Edition\vs_$Edition.exe"
$DownloadLink = $UniqURLs[$listboxEd.SelectedIndex]
$Prm = "--all --layout .\$Edition"
$PrmLng = " --lang"

if ($listboxLng.SelectedItems.Count -ne 14) {
    for ($i = 0; $i -lt $Languages[0].Count; $i++) {
        if ($listboxLng.GetSelected($i) -eq $true) {
            $PrmLng += $Languages[0][$i]
        }
    }
    $Prm += $PrmLng
}

Write-Host "Выбранная редакция Visual Studio: " -ForegroundColor Blue -NoNewline
Write-Host $Edition -ForegroundColor White
Write-Host "Получаем ссылку на файл онлайн-инсталлятора... " -ForegroundColor Blue -NoNewline

$result = Invoke-WebRequest $DownloadLink
if ($result.StatusCode -ne 200) {
    ShowErrorNExit($result.StatusCode)
    exit
}

$URL = Select-String -InputObject $result.Content -Pattern "http\S+\.exe"
if ($URL.Matches.Lenght -eq 0) {
    ShowErrorNExit(204)
    exit
}

Write-Host "${$URL.Matches.Value}`n"
Write-Host "Скачиваем онлайн-инсталлятор... " -ForegroundColor Blue -NoNewline
$result = Invoke-WebRequest $URL.Matches.Value -OutFile "installer.tmp"
Write-Host $OKText[$LngIndex]"`n" -ForegroundColor Green

if ((Test-Path -Path $Exe) -eq $false) {
    if ((Test-Path -Path $Edition) -eq $false) {
        $result = New-Item -Path $Edition -ItemType Directory
    }
    Move-Item -Path ".\installer.tmp" -Destination $Exe
}
else {
    $HashOld = Get-FileHash $Exe -Algorithm SHA1
    $HashNew = Get-FileHash "installer.tmp" -Algorithm SHA1
    if ($HashOld -ne $HashNew) {
        Remove-Item -Path $Exe
        Move-Item -Path "installer.tmp" -Destination $Exe
        
    }
    else {
        Remove-Item -Path "installer.tmp"
    }
}
Write-Host "Запуск онлайн-инсталлятора Visual Studio в режиме создания/обновления оффлайн-каталога... " -ForegroundColor Blue -NoNewline

$Process = Start-Process -FilePath $Exe -ArgumentList $Prm -Wait -PassThru
if ($Process.ExitCode -ne 0) {
    ShowErrorNExit($Process.ExitCode)
    exit
}
Write-Host $OKText[$LngIndex]"`n" -ForegroundColor Green
$Arch = ".\$Edition\Archive"
if ((Test-Path -Path $Arch) -eq $True) {
    $Cln = ""
    Get-ChildItem -Path $Arch -Directory | ForEach-Object $Cln = (($Cln -eq "") ? "" : $Cln + " ") + "--clean .\Archive\$_\Catalog.json"
    Write-Host "Запуск онлайн-инсталлятора Visual Studio в режиме очистки старых пакетов... " -ForegroundColor Blue -NoNewline
    $Process = Start-Process -FilePath $Exe -ArgumentList $Prm $Cln -Wait -PassThru
    if ($Process.ExitCode -ne 0) {
        ShowErrorNExit($Process.ExitCode)
        exit
    }
    Write-Host $OKText[$LngIndex]"`n" -ForegroundColor Green
    Remove-Item -Path $Arch -Recurse
    Write-Host "Удаление служебных файлов завершено.`nОбновление Visual Studio завершено.`n" -ForegroundColor Blue
    exit
}