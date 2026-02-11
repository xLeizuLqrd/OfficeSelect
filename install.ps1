[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Установщик Microsoft Office"

function Show-ModeMenu {
    Clear-Host
    Write-Host "        :::    :::       :::        ::::::::::       :::::::::::       :::::::::      :::    :::" -ForegroundColor Cyan
    Write-Host "       :+:    :+:       :+:        :+:                  :+:                :+:       :+:    :+:" -ForegroundColor Cyan
    Write-Host "       +:+  +:+        +:+        +:+                  +:+               +:+        +:+    +:+ " -ForegroundColor Cyan
    Write-Host "       +#++:+         +#+        +#++:++#             +#+              +#+         +#+    +:+  " -ForegroundColor Cyan
    Write-Host "     +#+  +#+        +#+        +#+                  +#+             +#+          +#+    +#+   " -ForegroundColor Cyan
    Write-Host "   #+#    #+#       #+#        #+#                  #+#            #+#           #+#    #+#    " -ForegroundColor Cyan
    Write-Host "  ###    ###       ########## ##########       ###########       #########       ########      " -ForegroundColor Cyan
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВЫБОР РЕЖИМА УСТАНОВКИ" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] Полная установка (удалит старый Office)" -ForegroundColor Yellow
    Write-Host "[2] Добавить программы к существующему Office" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    do {
        $mode = Read-Host "Выберите режим (1 или 2)"
        
        switch ($mode) {
            "1" {
                $script:RemoveMSI = $true
                $script:ModeName = "ПОЛНАЯ УСТАНОВКА"
                Show-MainMenu
                return
            }
            "2" {
                $script:RemoveMSI = $false
                $script:ModeName = "ДОБАВЛЕНИЕ ПРОГРАММ"
                Show-MainMenu
                return
            }
            default {
                Write-Host "Ошибка: Неверный выбор! Введите 1 или 2" -ForegroundColor Red
            }
        }
    } while ($true)
}

function Show-MainMenu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВЫБОР ПРОГРАММ MICROSOFT OFFICE" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Выберите цифры через пробел:" -ForegroundColor Gray
    Write-Host ""
    Write-Host "[1] Word       [2] Excel" -ForegroundColor Yellow
    Write-Host "[3] PowerPoint [4] Outlook" -ForegroundColor Yellow
    Write-Host "[5] Access     [6] Publisher" -ForegroundColor Yellow
    Write-Host "[7] OneNote    [8] OneDrive" -ForegroundColor Yellow
    Write-Host "[9] Teams      [10] Lync" -ForegroundColor Yellow
    Write-Host "[0] Назад" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Просто Enter - установить все программы" -ForegroundColor Gray
    Write-Host "РЕЖИМ: $($script:ModeName)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $input = Read-Host "Ваш выбор"
    
    if ($input -eq "0") {
        Show-ModeMenu
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($input)) {
        $script:InstallAll = $true
        $script:SelectedApps = @(1..10)
    } else {
        if ($input -match '^[0-9\s]+$') {
            $script:InstallAll = $false
            $script:SelectedApps = $input -split '\s+' | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le 10 }
            
            if ($script:SelectedApps.Count -eq 0) {
                Write-Host "Ошибка: Неверный ввод!" -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-MainMenu
                return
            }
        } else {
            Write-Host "Ошибка: Неверный формат ввода!" -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-MainMenu
            return
        }
    }
    
    Start-Installation
}

function Start-Installation {

    $workDir = "$env:TEMP\OfficeInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    
    try {
        $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19628-20192.exe"
        $odtPath = Join-Path $workDir "ODTSetup.exe"
        
        $progressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        $progressPreference = 'Continue'
        
        $extractDir = Join-Path $workDir "OfficeSetup"
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        Start-Process -FilePath $odtPath -ArgumentList "/extract:`"$extractDir`" /quiet" -Wait -NoNewWindow
        $setupPath = Join-Path $extractDir "setup.exe"
        
        $xmlContent = @()
        $xmlContent += '<?xml version="1.0" encoding="utf-8"?>'
        $xmlContent += '<Configuration>'
        $xmlContent += '  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">'
        $xmlContent += '    <Product ID="ProPlus2024Volume">'
        $xmlContent += '      <Language ID="ru-ru" />'
        
        if (-not $script:InstallAll) {
            $appMap = @{
                1 = "Word"; 2 = "Excel"; 3 = "PowerPoint"; 4 = "Outlook"
                5 = "Access"; 6 = "Publisher"; 7 = "OneNote"; 8 = "OneDrive"
                9 = "Teams"; 10 = "Lync"
            }
            $allApps = 1..10
            foreach ($appNum in $allApps) {
                if ($appNum -notin $script:SelectedApps) {
                    $xmlContent += "      <ExcludeApp ID=`"$($appMap[$appNum])`" />"
                }
            }
        }
        
        $xmlContent += '    </Product>'
        $xmlContent += '  </Add>'
        if ($script:RemoveMSI) { $xmlContent += '  <RemoveMSI />' }
        $xmlContent += '  <Display Level="None" AcceptEULA="TRUE" />'
        $xmlContent += '  <Property Name="AUTOACTIVATE" Value="1" />'
        $xmlContent += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
        $xmlContent += '</Configuration>'
        
        $xmlContent | Out-File -FilePath "$workDir\configuration.xml" -Encoding UTF8 -Force
        
        $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "onenote", "mspub")
        foreach ($process in $officeProcesses) {
            Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        }
        
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "     УСТАНОВКА MICROSOFT OFFICE" -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Установка Microsoft Office начата..." -ForegroundColor Green
        Write-Host ""
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $setupPath
        $psi.Arguments = "/configure `"$workDir\configuration.xml`""
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = $workDir
        
        $p = [System.Diagnostics.Process]::Start($psi)
        
        $totalSeconds = 0
        $maxSeconds = 1800
        
        while (-not $p.HasExited) {
            $totalSeconds++
            $percent = [math]::Min(100, [math]::Round(($totalSeconds / $maxSeconds) * 100))
            
            $bar = ""
            $barLength = 50
            $filled = [math]::Floor(($percent / 100) * $barLength)
            
            for ($i = 0; $i -lt $barLength; $i++) {
                if ($i -lt $filled) { $bar += "█" } else { $bar += "░" }
            }
            
            $minutes = [math]::Floor($totalSeconds / 60)
            $seconds = $totalSeconds % 60
            
            Write-Host "`rЗагрузка и установка: [$bar] $percent%  ${minutes}:$($seconds.ToString('00'))" -ForegroundColor Cyan -NoNewline
            
            Start-Sleep -Seconds 1
        }
        
        Write-Host ""
        Write-Host ""
        
        $exitCode = $p.ExitCode
        
        if ($exitCode -eq 0 -or $exitCode -eq 3010 -or $exitCode -eq 17002) {
            Write-Host "✓ Установка завершена успешно!" -ForegroundColor Green
        } else {
            Write-Host "✗ Ошибка установки (код: $exitCode)" -ForegroundColor Red
        }
        
    } catch {
        Write-Host ""
        Write-Host "✗ Ошибка: $_" -ForegroundColor Red
    } finally {
        try {
            Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] Вернуться в главное меню" -ForegroundColor Yellow
    Write-Host "[2] Выход" -ForegroundColor Yellow
    Write-Host ""
    
    do {
        $choice = Read-Host "Ваш выбор"
        switch ($choice) {
            "1" { Show-ModeMenu; return }
            "2" { exit }
            default { Write-Host "Неверный выбор. Введите 1 или 2" -ForegroundColor Red }
        }
    } while ($true)
}

$script:RemoveMSI = $false
$script:ModeName = ""
$script:InstallAll = $false
$script:SelectedApps = @()

Show-ModeMenu
