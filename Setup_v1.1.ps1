#region [0. 인코딩 강제 설정 - 한글 깨짐 방지]
$OutputEncoding = [System.Text.Encoding]::UTF8
#endregion

#region [1. 환경 설정 및 초기화]
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$configUrl = "https://raw.githubusercontent.com/limhyunji0620-max/qa-setup-config/refs/heads/main/config.json"
#endregion

#region [2. 데이터 로드]
try {
    $config = Invoke-RestMethod -Uri $configUrl -UseBasicParsing
} catch {
    [System.Windows.Forms.MessageBox]::Show("GitHub 연결 실패")
    exit
}
#endregion

#region [3. 메인 UI 디자인]
$form = New-Object System.Windows.Forms.Form
$form.Text = " QA Setup Launcher"
$form.Size = New-Object System.Drawing.Size(480, 650)
$form.BackColor = [System.Drawing.Color]::White
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false        
$form.TopMost = $false

# 폰트 정의
$fontTitle = New-Object System.Drawing.Font("맑은 고딕", 15, [System.Drawing.FontStyle]::Bold)
$fontList = New-Object System.Drawing.Font("맑은 고딕", 10)
$fontButton = New-Object System.Drawing.Font("맑은 고딕", 11, [System.Drawing.FontStyle]::Bold)

# 제목
$header = New-Object System.Windows.Forms.Label
$header.Text = "설치 패키지 선택"
$header.Location = New-Object System.Drawing.Point(25, 25)
$header.AutoSize = $true
$header.Font = $fontTitle
$header.ForeColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$form.Controls.Add($header)

# 전체선택 토글 버튼
$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "전체 선택"
$btnToggle.Location = New-Object System.Drawing.Point(340, 28)
$btnToggle.Size = New-Object System.Drawing.Size(90, 32)
$btnToggle.FlatStyle = "Flat"
$btnToggle.FlatAppearance.BorderSize = 1
$btnToggle.BackColor = [System.Drawing.Color]::White
$btnToggle.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnToggle)

# 개선된 체크리스트 박스
$checkedList = New-Object System.Windows.Forms.CheckedListBox
$checkedList.Location = New-Object System.Drawing.Point(25, 80)
$checkedList.Size = New-Object System.Drawing.Size(410, 380)
$checkedList.BorderStyle = "None"
$checkedList.CheckOnClick = $true
$checkedList.Font = $fontList
$checkedList.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$checkedList.ItemHeight = 35 

foreach ($pgm in $config.programs) { 
    [void]$checkedList.Items.Add($pgm.name) 
}
$form.Controls.Add($checkedList)

# 하단 상태 표시줄
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "준비 상태"
$statusLabel.Location = New-Object System.Drawing.Point(25, 475)
$statusLabel.Size = New-Object System.Drawing.Size(410, 25)
$statusLabel.ForeColor = [System.Drawing.Color]::SteelBlue
$form.Controls.Add($statusLabel)

# 설치 시작 버튼
$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "프로그램 설치 시작"
$btnInstall.Location = New-Object System.Drawing.Point(25, 510)
$btnInstall.Size = New-Object System.Drawing.Size(410, 65)
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnInstall.ForeColor = [System.Drawing.Color]::White
$btnInstall.FlatStyle = "Flat"
$btnInstall.FlatAppearance.BorderSize = 0
$btnInstall.Font = $fontButton
$btnInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnInstall)
#endregion

#region [4. 로직 처리]

$script:allChecked = $false
$global:isCancelled = $false

# 1. 전체 선택 토글
$btnToggle.Add_Click({
    $script:allChecked = !$script:allChecked
    for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
        $checkedList.SetItemChecked($i, $script:allChecked)
    }
    $btnToggle.Text = if ($script:allChecked) { "전체 해제" } else { "전체 선택" }
    $btnToggle.BackColor = if ($script:allChecked) { [System.Drawing.Color]::FromArgb(240, 240, 240) } else { [System.Drawing.Color]::White }
})

# 2. 설치 및 취소 통합 로직
$btnInstall.Add_Click({
    if ($btnInstall.Text -eq "설치 취소") {
        $cancelNotice = "설치를 중단하시겠습니까?`n`n1. 대기 중인 프로그램은 제외됩니다.`n2. 이미 실행된 설치창은 직접 닫아야 합니다."
        if ([System.Windows.Forms.MessageBox]::Show($cancelNotice, "설치 중단", "YesNo", "Warning") -eq "Yes") {
            $global:isCancelled = $true
            $statusLabel.Text = "중단 대기 중..."
            $btnInstall.Enabled = $false
        }
        return
    }

    if ($checkedList.CheckedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("설치할 프로그램을 선택해주세요.", "알림")
        return
    }

    # UI 상태 변경
    $global:isCancelled = $false
    $checkedList.Enabled = $btnToggle.Enabled = $false
    $btnInstall.Text = "설치 취소"; $btnInstall.BackColor = [System.Drawing.Color]::Firebrick
    
    $webClient = New-Object System.Net.WebClient
    $selectedItems = @($checkedList.CheckedItems)

    foreach ($itemName in $selectedItems) {
        if ($global:isCancelled) { break }

        $target = $config.programs | Where-Object { $_.name -eq $itemName }
        $random = Get-Random -Minimum 1000 -Maximum 9999
        $isZip = $target.url.Contains(".zip")
        $extension = if ($isZip) { ".zip" } else { ".exe" }
        
        $tempPath = Join-Path $env:TEMP "$($target.id)_$random$extension"
        $unzipPath = Join-Path $env:TEMP "$($target.id)_dir_$random"

        $statusLabel.Text = "● 다운로드 중: $($target.name)"
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # 다운로드
            $webClient.DownloadFileAsync($target.url, $tempPath)
            while ($webClient.IsBusy) {
                if ($global:isCancelled) { $webClient.CancelAsync(); break }
                [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
            }
            if ($global:isCancelled) { break }

            $executablePath = $tempPath

            # ZIP 처리
            if ($isZip) {
                $statusLabel.Text = "● 압축 해제 중: $($target.name)"
                [System.Windows.Forms.Application]::DoEvents()
                Expand-Archive -Path $tempPath -DestinationPath $unzipPath -Force
                $executablePath = Get-ChildItem -Path $unzipPath -Include *.msi, *.exe -Recurse | Select-Object -First 1 -ExpandProperty FullName
            }

            # --- 특수 설치 로직 분기 ---
            if ($target.id -eq "ms_office") {
                # [MS Office ODT 특화 로직]
                $statusLabel.Text = "● Office 설치 구성 중..."
                # 1. ODT 추출기 실행 (압축해제 역할)
                Start-Process -FilePath $tempPath -ArgumentList "/extract:`"$unzipPath`" /quiet" -Wait
                
                # 2. configuration.xml 생성
                $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="ko-kr"/>
    </Product>
  </Add>
</Configuration>
"@
                $xmlPath = Join-Path $unzipPath "configuration.xml"
                $xmlContent | Out-File -FilePath $xmlPath -Encoding utf8

                # 3. Setup.exe 실행
                $setupExe = Join-Path $unzipPath "setup.exe"
                if (Test-Path $setupExe) {
                    $statusLabel.Text = "● Office 설치 진행 중 (화면 변화가 없을 수 있음)..."
                    Start-Process -FilePath $setupExe -ArgumentList "/configure `"$xmlPath`"" -Wait
                }
            }
            elseif ($target.id -eq "capture_auto") {
                # [로컬 복사 및 바로가기 로직]
                $installDir = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "QA_Tools\$($target.id)"
                if (!(Test-Path $installDir)) { New-Item -Path $installDir -ItemType Directory -Force }
                Copy-Item -Path "$unzipPath\*" -Destination $installDir -Recurse -Force
                
                $exeFile = Get-ChildItem -Path $installDir -Filter "rawdata_capture_시스템_v0.1.exe" -Recurse | Select-Object -First 1
                if ($exeFile) {
                    $wsh = New-Object -ComObject WScript.Shell
                    $shortcut = $wsh.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "rawdata_capture.lnk"))
                    $shortcut.TargetPath = $exeFile.FullName
                    $shortcut.WorkingDirectory = $exeFile.DirectoryName
                    $shortcut.Save()
                }
            }
            else {
                # [일반 설치 로직]
                if ($executablePath) {
                    $statusLabel.Text = "● 설치 진행 중: $($target.name)"
                    $process = Start-Process -FilePath $executablePath -ArgumentList $target.args -PassThru
                    while (!$process.HasExited) {
                        if ($global:isCancelled) { break }
                        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
                    }
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("$($target.name) 오류: $($_.Exception.Message)")
        } finally {
            # 파일 정리
            if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
            if (Test-Path $unzipPath) { Remove-Item $unzipPath -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }

    # UI 복구 루틴
    $btnInstall.Text = "프로그램 설치 시작"
    $btnInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnInstall.Enabled = $checkedList.Enabled = $btnToggle.Enabled = $true
    $webClient.Dispose()

    if ($global:isCancelled) {
        [System.Windows.Forms.MessageBox]::Show("설치가 중단되었습니다.", "알림")
        $statusLabel.Text = "설치 중단됨"
    } else {
        $statusLabel.Text = "모든 작업 완료"
        [System.Windows.Forms.MessageBox]::Show("모든 설치가 완료되었습니다.", "완료")
    }
    $statusLabel.Text = "준비 상태"
})
#endregion

#region [5. 실행]
[void]$form.ShowDialog()
#endregion