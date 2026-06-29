@echo off
REM Builds a standalone outlook_attachment_downloader.exe
REM Run this ON WINDOWS, in the same folder as outlook_attachment_downloader.py
REM (PyInstaller does not cross-compile -- this must be built on Windows).

py -m pip install --upgrade pyinstaller pywin32

REM --hidden-import win32timezone: PyInstaller doesn't auto-detect this one,
REM and the script touches COM datetime properties (ReceivedTime, SentOn)
REM that pull it in at runtime. Without it, the frozen exe throws
REM "ModuleNotFoundError: No module named 'win32timezone'" the first time
REM it reads a date from an email.
py -m PyInstaller --onefile --name outlook_attachment_downloader --hidden-import win32timezone outlook_attachment_downloader.py

echo.
echo Done. The standalone executable is at dist\outlook_attachment_downloader.exe
echo That single file is what you hand to other users -- they still need
echo Outlook desktop installed and signed in, just not Python.
pause