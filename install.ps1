[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Установщик Microsoft Office"

function Show-ModeMenu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВЫБОР РЕЖИМА УСТАНОВКИ" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] Полная установка (удалит старый MSI Office)" -ForegroundColor Yellow
    Write-Host "[2] Добавить программы к существующему Office" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    do {
        $mode = Read-Host "Выберите режим (1 или 2)"
        switch ($mode) {
            "1" { $script:RemoveMSI = $true; $script:ModeName = "ПОЛНАЯ"; Show-MainMenu; return }
            "2" { $script:RemoveMSI = $false; $script:ModeName = "ДОБАВЛЕНИЕ"; Show-MainMenu; return }
            default { Write-Host "Ошибка! Введите 1 или 2" -ForegroundColor Red }
        }
    } while ($true)
}

function Show-MainMenu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВЫБОР ПРОГРАММ" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] Word       [2] Excel" -ForegroundColor Yellow
    Write-Host "[3] PowerPoint [4] Outlook" -ForegroundColor Yellow
    Write-Host "[5] Access     [6] Publisher" -ForegroundColor Yellow
    Write-Host "[7] OneNote    [8] OneDrive" -ForegroundColor Yellow
    Write-Host "[9] Teams      [10] Lync" -ForegroundColor Yellow
    Write-Host "[0] Назад" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Enter - все программы" -ForegroundColor Gray
    Write-Host "РЕЖИМ: $($script:ModeName)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $input = Read-Host "Ваш выбор"
    if ($input -eq "0") { Show-ModeMenu; return }
    
    if ([string]::IsNullOrWhiteSpace($input)) {
        $script:InstallAll = $true
        $script:SelectedApps = @(1..10)
    } else {
        if ($input -match '^[0-9\s]+$') {
            $script:InstallAll = $false
            $script:SelectedApps = $input -split '\s+' | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le 10 }
            if ($script:SelectedApps.Count -eq 0) { Write-Host "Ошибка!" -ForegroundColor Red; Start-Sleep 1; Show-MainMenu; return }
        } else { Write-Host "Ошибка!" -ForegroundColor Red; Start-Sleep 1; Show-MainMenu; return }
    }
    
    Start-Installation
}

function Start-Installation {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     УСТАНОВКА OFFICE LTSC 2024" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $workDir = "$env:TEMP\OfficeInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    
    try {
        $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19628-20192.exe"
        $odtPath = Join-Path $workDir "ODTSetup.exe"
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        
        $extractDir = Join-Path $workDir "OfficeSetup"
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        Start-Process -FilePath $odtPath -ArgumentList "/extract:`"$extractDir`" /quiet" -Wait -NoNewWindow
        $setupPath = Join-Path $extractDir "setup.exe"
        
        $appMap = @{1="Word";2="Excel";3="PowerPoint";4="Outlook";5="Access";6="Publisher";7="OneNote";8="OneDrive";9="Teams";10="Lync"}
        
        $xmlContent = @()
        $xmlContent += '<?xml version="1.0" encoding="utf-8"?>'
        $xmlContent += '<Configuration>'
        $xmlContent += '  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">'
        $xmlContent += '    <Product ID="ProPlus2024Volume">'
        $xmlContent += '      <Language ID="ru-ru" />'
        $xmlContent += '      <ExcludeApp ID="Groove" />'
        $xmlContent += '      <ExcludeApp ID="Bing" />'
        
        if (-not $script:InstallAll) {
            foreach ($appNum in 1..10) {
                if ($appNum -notin $script:SelectedApps) {
                    $xmlContent += "      <ExcludeApp ID=`"$($appMap[$appNum])`" />"
                }
            }
        }
        
        $xmlContent += '    </Product>'
        $xmlContent += '  </Add>'
        
        if ($script:RemoveMSI) {
            $xmlContent += '  <RemoveMSI />'
        }
        
        $xmlContent += '  <MigrateArchitecture>TRUE</MigrateArchitecture>'
        $xmlContent += '  <Property Name="SharedComputerLicensing" Value="0" />'
        $xmlContent += '  <Display Level="None" AcceptEULA="TRUE" />'
        $xmlContent += '  <Property Name="AUTOACTIVATE" Value="1" />'
        $xmlContent += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
        $xmlContent += '</Configuration>'
        
        $xmlContent | Out-File -FilePath "$workDir\configuration.xml" -Encoding UTF8 -Force
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $setupPath
        $psi.Arguments = "/configure `"$workDir\configuration.xml`""
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = $workDir
        
        $p = [System.Diagnostics.Process]::Start($psi)
        
        $barLength = 40
        $targetSize = 0
        $downloaded = 0
        
        while (-not $p.HasExited) {
            $cabFiles = Get-ChildItem "$env:TEMP\*.cab" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "stream" -and $_.Length -gt 0 }
            
            if ($cabFiles.Count -gt 0 -and $targetSize -eq 0) {
                foreach ($file in $cabFiles) {
                    try {
                        $content = [System.IO.File]::ReadAllText($file.FullName) -replace "`0", ""
                        if ($content -match "url=(.*?\.cab)") {
                            $head = Invoke-WebRequest -Uri $matches[1] -Method Head -UseBasicParsing -ErrorAction SilentlyContinue
                            if ($head.Headers.'Content-Length') {
                                $targetSize = [int]$head.Headers.'Content-Length'
                                break
                            }
                        }
                    } catch {}
                }
            }
            
            $downloaded = 0
            $cabFiles = Get-ChildItem "$env:TEMP\*.cab" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "stream" }
            foreach ($file in $cabFiles) {
                $downloaded += $file.Length
            }
            
            if ($targetSize -gt 0 -and $downloaded -gt 0) {
                $percent = [math]::Min(99, [math]::Round(($downloaded / $targetSize) * 100))
                $filled = [math]::Floor(($percent / 100) * $barLength)
                $bar = ""
                for ($i = 0; $i -lt $barLength; $i++) {
                    if ($i -lt $filled) { $bar += "█" } else { $bar += "░" }
                }
                
                $downloadedMB = [math]::Round($downloaded / 1MB, 1)
                $totalMB = [math]::Round($targetSize / 1MB, 1)
                Write-Host "`rЗагрузка: [$bar] $percent%  ${downloadedMB}MB/${totalMB}MB" -ForegroundColor Cyan -NoNewline
            } else {
                Write-Host "`rЗагрузка: [$('░'*$barLength)] 0%   0MB/??MB" -ForegroundColor Cyan -NoNewline
            }
            
            Start-Sleep -Milliseconds 500
        }
        
        $exitCode = $p.ExitCode
        
        if ($targetSize -gt 0) {
            Write-Host "`rЗагрузка: [$('█'*$barLength)] 100%  $([math]::Round($downloaded/1MB,1))MB/$([math]::Round($targetSize/1MB,1))MB" -ForegroundColor Green
        } else {
            Write-Host "`rЗагрузка: [$('█'*$barLength)] 100%  " -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host ""
        
        if ($exitCode -eq 0 -or $exitCode -eq 3010 -or $exitCode -eq 17002) {
            Write-Host "✅ Установка Office LTSC 2024 завершена!" -ForegroundColor Green
        } else {
            Write-Host "❌ Ошибка установки (код: $exitCode)" -ForegroundColor Red
        }
        
    } catch {
        Write-Host ""
        Write-Host "❌ Ошибка: $_" -ForegroundColor Red
    } finally {
        Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] В главное меню" -ForegroundColor Yellow
    Write-Host "[2] Выход" -ForegroundColor Yellow
    Write-Host ""
    
    do {
        $choice = Read-Host "Ваш выбор"
        switch ($choice) {
            "1" { Show-ModeMenu; return }
            "2" { exit }
            default { Write-Host "Введите 1 или 2" -ForegroundColor Red }
        }
    } while ($true)
}

$script:RemoveMSI = $false
$script:ModeName = ""
$script:InstallAll = $false
$script:SelectedApps = @()

Show-ModeMenu
