cd "C:\Users\Hankock\PyInstallerEXE\Tool"
set folder="C:\Users\Hankock\PyInstallerEXE\Tool"
cd /d %folder%
for /F "delims=" %%i in ('dir /b') do (rmdir "%%i" /s/q || del "%%i" /s/q)
pyinstaller --onefile ^
"C:\Users\Hankock\workspace\XML_Report\Source\main.py"
pause