Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
function ShowErrorNExit($ErrorCode) {
    Write-Host $ErrorText[$LngIndex] -ForegroundColor Red -NoNewline
    " 0x{0:X} ({0:D})`n" -f $ErrorCode
    Write-Host $PaKText[$LngIndex]"`n"
    [void] $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC")
}

$Languages = @(@("en-US", "es-ES", "it-IT", "zh-CN", "zh-TW", "ko-KR", "de-DE", "pl-PL", "pt-BR", "ru-RU", "tr-TR", "fr-FR", "cs-CZ", "ja-JP"), @())
$OKText = @("OK", "Aceptar", "OK", "确定", "確定", "확인", "OK", "OK", "OK", "OK", "Tamam", "OK", "OK", "OK")
$CancelText = @("Cancel", "Cancelar", "Annulla", "取消", "取消", "취소", "Abbrechen", "Anuluj", "Cancelar", "Отмена", "İptal", "Annuler", "Storno", "キャンセル")
$WindowTitle = @("Offline package settings",
                 "Configuración de paquetes sin conexión",
                 "Impostazioni del pacchetto offline",
                 "离线包设置",
                 "離線包設置",
                 "오프라인 패키지 설정",
                 "Offline-Paketeinstellungen",
                 "Ustawienia pakietu offline",
                 "Configurações de pacote off-line",
                 "Настройки автономного пакета",
                 "Çevrimdışı paket ayarları",
                 "Paramètres du package hors ligne",
                 "Nastavení offline balíku",
                 "オフラインパッケージの設定")
$EditionTitle = @("Visual Studio Edition:",
                  "Edición de Visual Studio:",
                  "Edizione di Visual Studio:",
                  "Visual Studio 版本:",
                  "Visual Studio 的版本:",
                  "Visual Studio 버전:",
                  "Visual Studio-Edition:",
                  "Wersja Visual Studio:",
                  "Edição do Visual Studio:",
                  "Редакция Visual Studio:",
                  "Visual Studio sürümü:",
                  "Édition de Visual Studio:",
                  "Edice Visual Studio:",
                  "Visual Studio エディション:")
$LangTitle = @("Needed languages:",
               "Seleccione los idiomas necesarios:",
               "Seleziona le lingue necessarie:",
               "选择需要的语言：",
               "選擇需要的語言：",
               "필요한 언어 선택 :",
               "Gewünschte sprachen auswählen:",
               "Wybierz potrzebne języki:",
               "Selecione os idiomas necessários:",
               "Выберите необходимые языки:",
               "Gerekli dilleri seçin:",
               "Sélectionnez les langues nécessaires :",
               "Vyberte potřebné jazyky:",
               "必要な言語を選択してください:")
$ErrorText = @("Error", "Error", "Errore", "错误", "錯誤", "오류", "Error", "Błąd", "Erro", "Ошибка", "Hata", "Erreur", "Chyba", "エラー")
$PaKText = @("Press any key...",
             "Presiona cualquier tecla...",
             "Premere un tasto qualsiasi...",
             "按任意键...",
             "按任意鍵...",
             "아무 키나 누르시오...",
             "Drücke irgendeine taste...",
             "Naciśnij dowolny klawisz...",
             "Pressione qualquer tecla...",
             "Нажмите любую клавишу...",
             "Herhangi bir tuşa basın...",
             "Appuyez sur une touche...",
             "Stiskněte libovolnou klávesu...",
             "どれかキーを押してください...")
$VersionText = @("PowerShell version must be 7.0 and higher",
                 "La versión de PowerShell debe ser 7.0 y superior",
                 "La versione di PowerShell deve essere 7.0 e successive",
                 "PowerShell 版本必须是 7.0 及更高版本",
                 "PowerShell 版本必須是 7.0 及更高版本",
                 "PowerShell 버전은 7.0 이상이어야 합니다.",
                 "PowerShell-Version muss 7.0 und höher sein",
                 "Wersja PowerShell musi być 7.0 lub nowsza",
                 "A versão do PowerShell deve ser 7.0 e superior",
                 "Версия PowerShell должна быть 7.0 и выше",
                 "PowerShell sürümü 7.0 ve üzeri olmalıdır",
                 "La version de PowerShell doit être 7.0 et supérieure",
                 "Verze prostředí PowerShell musí být 7.0 a vyšší",
                 "PowerShellのバージョンは7.0以降である必要があります")
$ConsoleText1 = @("Selected Visual Studio Edition: ",
                  "Edición de Visual Studio seleccionada: ",
                  "Edizione di Visual Studio selezionata: ",
                  "选定的 Visual Studio 版本：",
                  "選定的 Visual Studio 版本：",
                  "선택한 Visual Studio 에디션: ",
                  "Ausgewählte Visual Studio-Edition: ",
                  "Wybrana edycja programu Visual Studio: ",
                  "Edição selecionada do Visual Studio: ",
                  "Выбранная редакция Visual Studio: ",
                  "Seçilen Visual Studio sürümü: ",
                  "Édition Visual Studio sélectionnée: ",
                  "Vybraná edice Visual Studio: ",
                  "選択された Visual Studio エディション: ")
$ConsoleText2 = @("Getting a link to the online installer... ",
                  "Obteniendo un enlace al instalador en línea... ",
                  "Ottenere un collegamento al programma di installazione online... ",
                  "获取在线安装程序的链接... ",
                  "獲取在線安裝程序的鏈接...",
                  "온라인 설치 프로그램에 대한 링크를 가져 오는 중 ... ",
                  "Link zum online-Installer erhalten... ",
                  "Uzyskiwanie łącza do instalatora online... ",
                  "Obtendo um link para o instalador online... ",
                  "Получаем ссылку на файл онлайн-инсталлятора... ",
                  "Çevrimiçi yükleyiciye bir bağlantı alınıyor... ",
                  "Obtenir un lien vers le programme d'installation en ligne...",
                  "Získání odkazu na online instalační program... ",
                  "オンライン インストーラーへのリンクを取得しています... ")
$ConsoleText3 = @("Downloading online-installer... ",
                  "Descargando el instalador en línea... ",
                  "Download del programma di installazione online in corso... ",
                  "正在下载在线安装程序... ",
                  "正在下載在線安裝程序... ",
                  "온라인 설치 프로그램 다운로드 중 ... ",
                  "Online-Installationsprogramm wird heruntergeladen... ",
                  "Pobieranie instalatora online... ",
                  "Baixando o instalador online... ",
                  "Скачиваем онлайн-инсталлятор... ",
                  "Çevrimiçi yükleyici indiriliyor... ",
                  "Téléchargement du programme d'installation en ligne... ",
                  "Stahování instalačního programu online... ",
                  "オンライン インストーラーをダウンロードしています... ")
$ConsoleText4 = @("Launch the online installer in the offline catalog creation/update mode... ",
                  "Inicie el instalador en línea para crear / actualizar paquetes sin conexión ... ",
                  "Avvia il programma di installazione online per creare/aggiornare i pacchetti offline... ",
                  "为创建/更新离线包启动在线安装程序... ",
                  "為創建/更新離線包啟動在線安裝程序... ",
                  "오프라인 패키지 생성 / 업데이트를위한 온라인 설치 프로그램을 시작합니다 ... ",
                  "Starten Sie das online-Installationsprogramm für die offline-pakete erstellen/aktualisieren... ",
                  "Uruchom instalator online dla tworzenia/aktualizacji pakietów offline... ",
                  "Inicie o instalador online para criar / atualizar pacotes offline ... ",
                  "Запуск онлайн-инсталлятора в режиме загрузки/обновления оффлайн-пакетов... ",
                  "Çevrimdışı paketler oluşturmak/güncellemek için çevrimiçi yükleyiciyi başlatın... ",
                  "Lancez le programme d'installation en ligne pour créer/mettre à jour les packages hors ligne... ",
                  "Spusťte online instalační program pro vytváření / aktualizaci offline balíčků... ",
                  "オフライン パッケージを作成/更新するためのオンライン インストーラーを起動します... ")
$ConsoleText5 = @("Running the online installer in the cleanup mode from the obsolete packages... ",
                  "Ejecutando el instalador en línea en el modo de limpieza de los paquetes obsoletos... ",
                  "Esecuzione del programma di installazione online in modalità di pulizia dai pacchetti obsoleti... ",
                  "从过时的包中以清理模式运行在线安装程序... ",
                  "從過時的包中以清理模式運行在線安裝程序... ",
                  "오래된 패키지에서 정리 모드로 온라인 설치 프로그램 실행 ... ",
                  "Ausführen des Online-Installationsprogramms im bereinigungsmodus von den veralteten paketen... ",
                  "Uruchamianie instalatora online w trybie czyszczenia z przestarzałych pakietów... ",
                  "Executando o instalador online no modo de limpeza dos pacotes obsoletos... ",
                  "Запуск онлайн-инсталлятора в режиме очистки старых пакетов... ",
                  "Çevrimiçi yükleyiciyi eski paketlerden temizleme modunda çalıştırmak... ",
                  "Exécution du programme d'installation en ligne en mode nettoyage à partir des packages obsolètes... ",
                  "Spuštění online instalačního programu v režimu čištění ze zastaralých balíčků... ",
                  "廃止されたパッケージからクリーンアップ モードでオンライン インストーラーを実行しています... ")
$ConsoleText6 = @("Old packages removed.`nVisual Studio update completed.",
                  "Se eliminaron los paquetes antiguos.`nSe completó la actualización de Visual Studio.",
                  "Vecchi pacchetti rimossi.`nAggiornamento di Visual Studio completato.",
                  "旧包已删除。`nVisual Studio 更新已完成。",
                  "舊包已刪除。`nVisual Studio 更新已完成。",
                  "이전 패키지가 제거되었습니다.`nVisual Studio 업데이트가 완료되었습니다.",
                  "Alte Pakete entfernt.`nVisual Studio-Update abgeschlossen.",
                  "Usunięto stare pakiety.`nAktualizacja programu Visual Studio zakończona.",
                  "Pacotes antigos removidos.`nAtualização do Visual Studio concluída.",
                  "Старые пакеты удалены.`nОбновление Visual Studio завершено.",
                  "Eski paketler kaldırıldı.`nVisual Studio güncellemesi tamamlandı.",
                  "Anciens packages supprimés.`nMise à jour de Visual Studio terminée.",
                  "Staré balíčky odstraněny.`nAktualizace sady Visual Studio dokončena.",
                  "古いパッケージが削除されました。`nVisual Studio の更新が完了しました。")
$ConsoleText7 = @("No archives were found for the cleanup. Exiting immediately.",
                  "No se encontraron archivos para la limpieza. Saliendo de inmediato.",
                  "Nessun archivio trovato per la pulizia. Uscita immediata.",
                  "没有找到用于清理的档案。 立即退出。",
                  "沒有找到用於清理的檔案。 立即退出。",
                  "정리에 대한 아카이브를 찾을 수 없습니다. 즉시 종료합니다.",
                  "Für die Bereinigung wurden keine Archive gefunden. Ausstieg sofort.",
                  "Nie znaleziono archiwów do czyszczenia. Wyjście natychmiast.",
                  "Nenhum arquivo foi encontrado para a limpeza. Saindo imediatamente.",
                  "Архивы для очистки не найдены. Завершаем работу.",
                  "Temizleme için arşiv bulunamadı. Hemen çıkış.",
                  "Aucune archive n'a été trouvée pour le nettoyage. Sortie immédiate.",
                  "Pro vyčištění nebyly nalezeny žádné archivy. Okamžitý odchod.",
                  "クリーンアップ用のアーカイブは見つかりませんでした。 すぐに終了します。")

# Get all languages and filter with our list
$Cult = Get-Culture -ListAvailable
$LngName = $(Get-UICulture).Name
$LngIndex = $Languages[0].IndexOf($LngName)
if ($LngIndex -eq -1) {
    $LngIndex = 0
}
# Get the language readable name from the short form
foreach ($item in $Languages[0]) {
    $result = $Cult | Where-Object { $_.Name -like $item }
    $Languages[1] += (Get-Culture).TextInfo.ToTitleCase($result.DisplayName)
}

if ($Host.Version.Major -lt 7) {
    Write-Host $VersionText[$LngIndex]
    Exit
}

$ProgressPreference = 'SilentlyContinue'
# Get the web page with download links, take all unique URLs by mask
$Editions = @()
$AllURLs = Invoke-WebRequest "https://visualstudio.microsoft.com/$LngName/downloads/" | Select-String -InputObject { $_.Content } -Pattern 'http\S+studio\/\?sku\S+(?=\")' -AllMatches
$UniqURLs = $AllURLs.Matches.Value | Select-Object -Unique
foreach ($item in $UniqURLs) {
    $result = Select-String -InputObject $item -Pattern '(?<=sku\=)\w+'
    $Editions += $result.Matches.Value
}
<#
-------------------- Window form --------------------
#>
$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle[$LngIndex]                                        # Offline package settings
$form.Size = New-Object System.Drawing.Size(410,380)
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 20)
$label1.Size = New-Object System.Drawing.Size(150, 20)
$label1.Text = $EditionTitle[$LngIndex]                                     # Visual Studio Edition
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
$label2.Text = $LangTitle[$LngIndex]                                        # Needed languages
$form.Controls.Add($label2)

$listboxLng = New-Object System.Windows.Forms.ListBox
$listboxLng.Location = New-Object System.Drawing.Point(20, 70)
$listboxLng.Size = New-Object System.Drawing.Size(220, 20)
foreach ($lng in $Languages[1]) {
    [void] $listboxLng.Items.Add($lng)
}
$listboxLng.SetSelected($LngIndex, $true);
$listboxLng.SelectionMode = 'MultiExtended'
$listboxLng.Height = 260
$form.Controls.Add($listboxLng)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(265, 245)
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Text = $OKText[$LngIndex]                                         # OK button
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(265, 290)
$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
$cancelButton.Text = $CancelText[$LngIndex]                                 # Cancel button
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
$Exe = ".\vs_$Edition.exe"
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

Write-Host $ConsoleText1[$LngIndex] -ForegroundColor Cyan -NoNewline        # Selected Visual Studio Edition
Write-Host $Edition -ForegroundColor White
Write-Host $ConsoleText2[$LngIndex] -ForegroundColor Cyan -NoNewline        # Getting a link to the online installer

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
Write-Host $ConsoleText3[$LngIndex] -ForegroundColor Cyan -NoNewline        # Downloading online-installer
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
Write-Host $ConsoleText4[$LngIndex] -ForegroundColor Cyan -NoNewline        # Launch the online installer in the offline catalog creation/update mode

$Process = Start-Process -FilePath $Exe -ArgumentList $Prm -Wait -PassThru
if ($Process.ExitCode -ne 0) {
    ShowErrorNExit($Process.ExitCode)
    exit
}
Write-Host $OKText[$LngIndex]"`n" -ForegroundColor Green

Write-Host $ConsoleText5[$LngIndex] -ForegroundColor Cyan -NoNewline        # Running the online installer in the cleanup mode from the obsolete packages
$Arch = ".\$Edition\Archive"
if ((Test-Path -Path $Arch) -eq $True) {
    Get-ChildItem -Path $Arch -Directory | ForEach-Object -Begin { $Cln = "" } -Process { $Cln = (($Cln -eq "") ? "" : $Cln + " ") + "--clean .\Archive\$_\Catalog.json" }
    $Process = Start-Process -FilePath $Exe -ArgumentList $Prm $Cln -Wait -PassThru
    if ($Process.ExitCode -ne 0) {
        ShowErrorNExit($Process.ExitCode)
        exit
    }
    Write-Host -ForegroundColor Green $OKText[$LngIndex]"`n"
    Remove-Item -Path $Arch -Recurse
    Write-Host -ForegroundColor Cyan $ConsoleText6[$LngIndex]"`n"           # Old packages removed
    exit
}
else {
    Write-Host $ConsoleText7[$LngIndex] -ForegroundColor Green              # No archive founded
}
Remove-Item -Path $Exe
$ProgressPreference = 'Continue'