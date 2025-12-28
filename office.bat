@echo off
chcp 65001
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Запуск от имени администратора...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"
title Установка Microsoft Office
setlocal enabledelayedexpansion

if not exist "setup.exe" (
    echo Ошибка: setup.exe не найден в текущей папке!
    pause
    exit /b 1
)

:main_menu
cls
echo        :::    :::       :::        ::::::::::       :::::::::::       :::::::::      :::    :::
echo       :+:    :+:       :+:        :+:                  :+:                :+:       :+:    :+:
echo       +:+  +:+        +:+        +:+                  +:+               +:+        +:+    +:+
echo       +#++:+         +#+        +#++:++#             +#+              +#+         +#+    +:+
echo     +#+  +#+        +#+        +#+                  +#+             +#+          +#+    +#+
echo   #+#    #+#       #+#        #+#                  #+#            #+#           #+#    #+#
echo  ###    ###       ########## ##########       ###########       #########       ########
echo ========================================
echo     ВЫБОР ПРОГРАММ MICROSOFT OFFICE
echo ========================================
echo.
echo Выберите цифры через пробел:
echo.
echo [1] Word
echo [2] Excel
echo [3] PowerPoint
echo [4] Outlook
echo [5] Access
echo [6] Publisher
echo [7] OneNote
echo [8] OneDrive
echo [9] Teams
echo [10] Groove
echo [11] Lync
echo.
echo Просто Enter - установить все программы
echo.
echo ========================================
echo.

set /p "user_input=Ваш выбор: "

reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "std::wstring|US" /f

echo ^<Configuration^> > configuration.xml
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^> >> configuration.xml
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="Y63J7-9RNDJ-GD3BV-BDKBP-HH966"^> >> configuration.xml
echo       ^<Language ID="ru-ru" /^> >> configuration.xml

if "%user_input%"=="" (
    echo         ^<!-- Все программы будут установлены --^> >> configuration.xml
    echo Выбраны все программы.
) else (
    set "exclude_list="
    
    set "selected_word=0"
    set "selected_excel=0"
    set "selected_powerpoint=0"
    set "selected_outlook=0"
    set "selected_access=0"
    set "selected_publisher=0"
    set "selected_onenote=0"
    set "selected_onedrive=0"
    set "selected_teams=0"
    set "selected_groove=0"
    set "selected_lync=0"
    
    for %%i in (%user_input%) do (
        if "%%i"=="1" set "selected_word=1"
        if "%%i"=="2" set "selected_excel=1"
        if "%%i"=="3" set "selected_powerpoint=1"
        if "%%i"=="4" set "selected_outlook=1"
        if "%%i"=="5" set "selected_access=1"
        if "%%i"=="6" set "selected_publisher=1"
        if "%%i"=="7" set "selected_onenote=1"
        if "%%i"=="8" set "selected_onedrive=1"
        if "%%i"=="9" set "selected_teams=1"
        if "%%i"=="10" set "selected_groove=1"
        if "%%i"=="11" set "selected_lync=1"
    )
    
    if !selected_word!==0 (
        echo         ^<ExcludeApp ID="Word" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Word, "
    )
    if !selected_excel!==0 (
        echo         ^<ExcludeApp ID="Excel" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Excel, "
    )
    if !selected_powerpoint!==0 (
        echo         ^<ExcludeApp ID="PowerPoint" /^> >> configuration.xml
        set "exclude_list=!exclude_list!PowerPoint, "
    )
    if !selected_outlook!==0 (
        echo         ^<ExcludeApp ID="Outlook" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Outlook, "
    )
    if !selected_access!==0 (
        echo         ^<ExcludeApp ID="Access" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Access, "
    )
    if !selected_publisher!==0 (
        echo         ^<ExcludeApp ID="Publisher" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Publisher, "
    )
    if !selected_onenote!==0 (
        echo         ^<ExcludeApp ID="OneNote" /^> >> configuration.xml
        set "exclude_list=!exclude_list!OneNote, "
    )
    if !selected_onedrive!==0 (
        echo         ^<ExcludeApp ID="OneDrive" /^> >> configuration.xml
        set "exclude_list=!exclude_list!OneDrive, "
    )
    if !selected_teams!==0 (
        echo         ^<ExcludeApp ID="Teams" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Teams, "
    )
    if !selected_groove!==0 (
        echo         ^<ExcludeApp ID="Groove" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Groove, "
    )
    if !selected_lync!==0 (
        echo         ^<ExcludeApp ID="Lync" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Lync, "
    )
    
    echo.
    echo Вы выбрали:
    if !selected_word!==1 echo   - Word
    if !selected_excel!==1 echo   - Excel
    if !selected_powerpoint!==1 echo   - PowerPoint
    if !selected_outlook!==1 echo   - Outlook
    if !selected_access!==1 echo   - Access
    if !selected_publisher!==1 echo   - Publisher
    if !selected_onenote!==1 echo   - OneNote
    if !selected_onedrive!==1 echo   - OneDrive
    if !selected_teams!==1 echo   - Teams
    if !selected_groove!==1 echo   - Groove
    if !selected_lync!==1 echo   - Lync
    
    if "!exclude_list!" NEQ "" (
        echo.
        echo Не будут установлены: !exclude_list:~0,-2!
    )
)

echo     ^</Product^> >> configuration.xml
echo   ^</Add^> >> configuration.xml
echo   ^<RemoveMSI /^> >> configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> configuration.xml
echo ^</Configuration^> >> configuration.xml

echo.
echo ========================================
echo Конфигурация создана.
echo Запуск установки Office...
echo ========================================
echo.

setup.exe /configure configuration.xml

echo.
echo ========================================
echo Установка Office завершена!
echo ========================================
echo.
echo Microsoft Office успешно установлен!
echo.
echo Для начала работы откройте любое приложение Office.
echo.

:exit_script
echo ========================================
echo Процесс установки завершен!
echo ========================================
echo.

del configuration.xml >nul 2>&1

pause