@echo off
python -m pip install -r requirements.txt

REM Create a desktop shortcut
echo Creating desktop shortcut...
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Desktop\OpenOPen.lnk');$s.TargetPath='%~dp0OpenOPen.bat';$s.WorkingDirectory='%~dp0';$s.Save()"

python main.py
pause