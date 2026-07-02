@echo off
REM Builds a standalone outlook_attachment_downloader.exe with improved COM support
REM Run this ON WINDOWS, in the same folder as outlook_attachment_downloader.py
REM (PyInstaller does not cross-compile -- this must be built on Windows).

setlocal enabledelayedexpansion

py -m pip install --upgrade pyinstaller pywin32

REM Post-install hook: register COM if needed (for the build machine)
REM This ensures win32com can find COM type libraries
python -m PyInstaller.utils.win32 post_install_win32_service

REM --hidden-import flags: these modules are needed at runtime but PyInstaller
REM doesn't auto-detect them from static analysis:
REM   - win32timezone: used when reading ReceivedTime from COM objects
REM   - win32com.shell: may be needed for folder operations
REM   - pythoncom: COM initialization on startup
REM
REM We also use --collect-submodules to grab the entire win32com namespace,
REM ensuring all COM client stubs are available.

py -m PyInstaller ^
  --onefile ^
  --name outlook_attachment_downloader ^
  --hidden-import win32timezone ^
  --hidden-import win32com.shell ^
  --hidden-import pythoncom ^
  --collect-submodules win32com ^
  outlook_attachment_downloader.py

echo.
echo Done. The standalone executable is at dist\outlook_attachment_downloader.exe
echo That single file is what you hand to other users -- they still need
echo Outlook desktop installed and signed in, just not Python.
pause
