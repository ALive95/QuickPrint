# Builds dist\MAIN.exe using the MyEnv conda environment (where fitz/docx2pdf/pywin32 live).
# Usage: close MAIN.exe, then run  .\build.ps1
$E = "C:\Users\loryl\anaconda3\envs\MyEnv"
# Same PATH that "conda activate" sets, so PyInstaller finds the DLLs in Library\bin (tcl/tk, expat, ...)
$env:PATH = "$E;$E\Library\mingw-w64\bin;$E\Library\usr\bin;$E\Library\bin;$E\Scripts;$E\bin;" + $env:PATH
# Temporary build files go outside OneDrive to avoid file-lock errors
& "$E\python.exe" -m PyInstaller --noconsole --onefile --noconfirm --workpath "$env:TEMP\QuickPrint_build" --icon your_icon.ico MAIN.py
