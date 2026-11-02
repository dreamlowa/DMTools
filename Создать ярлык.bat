@echo off
setlocal
cd /d "%~dp0"

set "EXE=%CD%\delfin PDF_DM_CSV.exe"
if not exist "%EXE%" (
  echo Файл delfin PDF_DM_CSV.exe не найден рядом с этим скриптом.
  echo Убедитесь, что файлы программы распакованы из архива полностью.
  pause
  exit /b 1
)

set "SHORTCUT=%USERPROFILE%\Desktop\delfin PDF_DM_CSV.lnk"
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%EXE%'; $s.Description = 'delfin PDF_DM_CSV - извлечение кодов маркировки'; $s.Save()"

if exist "%SHORTCUT%" (
  echo Ярлык создан на рабочем столе: delfin PDF_DM_CSV
) else (
  set "SHORTCUT=%USERPROFILE%\OneDrive\Desktop\delfin PDF_DM_CSV.lnk"
  powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%EXE%'; $s.Description = 'delfin PDF_DM_CSV'; $s.Save()"
)
echo.
echo Готово! Ярлык добавлен на рабочий стол.
pause