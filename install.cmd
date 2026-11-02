@echo off
setlocal enabledelayedexpansion
title Установка delfin PDF_DM_CSV

echo ============================================
echo  Установка delfin PDF_DM_CSV
echo ============================================
echo.

rem Куда устанавливать (папка пользователя, без прав администратора)
set "APPDEST=%LOCALAPPDATA%\Programs\delfin PDF_DM_CSV"

rem Распаковка: этот скрипт лежит рядом с exe
set "SRCDIR=%~dp0"

if not exist "%SRCDIR%delfin PDF_DM_CSV.exe" (
  echo Не найден файл delfin PDF_DM_CSV.exe рядом с установщиком.
  echo Убедитесь, что архив распакован полностью.
  pause
  exit /b 1
)

if not exist "%APPDEST%" mkdir "%APPDEST%"

echo Копирую файлы программы...
copy /Y "%SRCDIR%delfin PDF_DM_CSV.exe" "%APPDEST%\delfin PDF_DM_CSV.exe" >nul
if exist "%SRCDIR%Читать.txt" copy /Y "%SRCDIR%Читать.txt" "%APPDEST%\Читать.txt" >nul
if exist "%SRCDIR%centrdelfin.ru logo9.svg" copy /Y "%SRCDIR%centrdelfin.ru logo9.svg" "%APPDEST%\" >nul
if exist "%SRCDIR%centrdelfin.ru logo13.svg" copy /Y "%SRCDIR%centrdelfin.ru logo13.svg" "%APPDEST%\" >nul

echo Создаю ярлык на рабочем столе...
set "SHORTCUT=%USERPROFILE%\Desktop\delfin PDF_DM_CSV.lnk"
powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath='%APPDEST%\delfin PDF_DM_CSV.exe'; $s.WorkingDirectory='%APPDEST%'; $s.Description='delfin PDF_DM_CSV - извлечение кодов маркировки'; $s.IconLocation='%APPDEST%\delfin PDF_DM_CSV.exe,0'; $s.Save()"

if not exist "%SHORTCUT%" (
  set "SHORTCUT=%USERPROFILE%\OneDrive\Desktop\delfin PDF_DM_CSV.lnk"
  powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath='%APPDEST%\delfin PDF_DM_CSV.exe'; $s.WorkingDirectory='%APPDEST%'; $s.Description='delfin PDF_DM_CSV'; $s.IconLocation='%APPDEST%\delfin PDF_DM_CSV.exe,0'; $s.Save()"
)

echo Создаю ярлык в меню Пуск...
set "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
if not exist "%STARTMENU%" mkdir "%STARTMENU%"
powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%STARTMENU%\delfin PDF_DM_CSV.lnk'); $s.TargetPath='%APPDEST%\delfin PDF_DM_CSV.exe'; $s.WorkingDirectory='%APPDEST%'; $s.Description='delfin PDF_DM_CSV'; $s.IconLocation='%APPDEST%\delfin PDF_DM_CSV.exe,0'; $s.Save()"

echo.
echo Установка завершена!
echo Программа установлена в: %APPDEST%
echo На рабочем столе и в меню Пуск добавлен ярлык.
echo.
echo Запустить программу сейчас?
choice /C YN /M "Запустить delfin PDF_DM_CSV (Y - да, N - нет)"
if "%ERRORLEVEL%"=="1" start "" "%APPDEST%\delfin PDF_DM_CSV.exe"

echo.
pause