@echo off
REM Builds a standalone outlook_attachment_downloader_gui.exe with improved COM support
REM Run this ON WINDOWS, with BOTH .py files in this same folder
REM (outlook_attachment_downloader.py is imported by the GUI script,
REM PyInstaller bundles it automatically -- the resulting .exe is fully
REM standalone, recipients only need that one file).
REM PyInstaller does not cross-compile -- this must be built on Windows.

setlocal enabledelayedexpansion

py -m pip install --upgrade pyinstaller pywin32

REM Post-install hook: register COM if needed (for the build machine)
REM This ensures win32com can find COM type libraries
python -m PyInstaller.utils.win32 post_install_win32_service

REM --hidden-import flags: these modules are needed at runtime but PyInstaller
REM doesn't auto-detect them from static analysis:
REM   - win32timezone: used when reading ReceivedTime from COM objects
REM   - win32com.shell: may be needed for folder operations
REM   - pythoncom: COM initialization in worker thread
REM
REM We also use --collect-submodules to grab the entire win32com namespace,
REM ensuring all COM client stubs are available.

py -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name outlook_attachment_downloader_gui ^
  --hidden-import win32timezone ^
  --hidden-import win32com.shell ^
  --hidden-import pythoncom ^
  --collect-submodules win32com ^
  outlook_attachment_downloader_gui.py

echo.
echo Done. The standalone executable is at dist\outlook_attachment_downloader_gui.exe
pause
