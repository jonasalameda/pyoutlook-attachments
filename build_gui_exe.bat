@echo off
REM Builds a standalone outlook_attachment_downloader_gui.exe
REM Run this ON WINDOWS, with BOTH .py files in this same folder
REM (outlook_attachment_downloader.py is imported by the GUI script,
REM PyInstaller bundles it automatically -- the resulting .exe is fully
REM standalone, recipients only need that one file).
REM PyInstaller does not cross-compile -- this must be built on Windows.

py -m pip install --upgrade pyinstaller pywin32

REM --windowed: no console box, since this is a real GUI now (unlike the
REM   CLI/REPL build, which keeps its console).
REM --hidden-import win32timezone: same COM-datetime gotcha as the CLI
REM   build -- this script also reads ReceivedTime via COM.
py -m PyInstaller --onefile --windowed --name outlook_attachment_downloader_gui --hidden-import win32timezone outlook_attachment_downloader_gui.py

echo.
echo Done. The standalone executable is at dist\outlook_attachment_downloader_gui.exe
pause
