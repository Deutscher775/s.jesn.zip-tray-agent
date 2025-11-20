<img src="https://s.jesn.zip/u/weasel/s.jesn.zip%20icon.png">

JesnZIP Tray Agent
===========================

This the tray agent that automatically uploads screenshots, pictures and videos inside your clipboard so https://s.jesn.zip file sharing.
## Note
This is completely loginless usable. A login is only required if you want to:
1. save files permanently
2. want them to show up in the "My  Files" on the webite for future uses

Max file size is 1Gb

# Installation
- Download the latest `installer.exe` in the releases tab
- Follow the instructions in the terminal

This will also:
- Ensure pip is up-to-date and install requirements from `requirements.txt`.
- Run PyInstaller to create a single-file (onefile) windowed executable using `ICON.ico` if present.
- Zip the created executable into a timestamped .zip located in the repository root.

Run it yourself (development or build the executable yourself)
-----------------------

```powershell
python .\JesnZIP-tray.py
```

Install runtime & build dependencies
------------------------------------

```powershell
python -m pip install -r .\requirements.txt
```



Notes
-----
- The tray agent uses Windows-only packages (`pywin32`, `win10toast`, `pystray`).
- The hard-coded upload endpoint in the tray script is `https://s.jesn.zip/api/upload`. Modify `UPLOAD_ENDPOINT` in `JesnZIP-tray.py` if needed.
- The autostart toggle creates/removes a shortcut in the user's Startup folder and should work for both running from script (python) and a packaged `.exe` built by PyInstaller.
- For very large video uploads you may need to increase the upload timeout or implement chunked upload.

Support
-------
If you need help running or building the tray agent, tell me what error you see and I'll help debug.



