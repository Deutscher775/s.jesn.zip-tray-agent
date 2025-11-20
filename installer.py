#!/usr/bin/env python3
"""
installer.py
Python helper to create a reproducible build for the JesnZIP tray agent.

This script replicates the behavior of the previous PowerShell helper:
- create (or reuse) a virtual environment in `.venv_build`
- upgrade pip/setuptools/wheel inside the venv
- install project requirements into the venv (per-package fallback; skip optional 'winrt')
- run PyInstaller from the venv to produce a single-file, windowed exe
- zip the produced executable into a timestamped archive
- optionally create a Start Menu shortcut

Run from the `TrayAgent` directory with the system Python you want to use for the build.
"""
from __future__ import annotations

import os
import sys
import subprocess
import shutil
import urllib.request
import venv
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime

if os.name == 'nt':
    import ctypes
    ctypes.windll.kernel32.SetConsoleTitleW("JesnZIP Tray Agent Installer")
else:
    raise NotImplementedError('This script currently only supports Windows builds.')

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / '.venv_build'
SCRIPT_NAME = 'JesnZIP-tray.py'
ICON_NAME = 'ICON.ico'
SETTINGS_NAME = 'tray_settings.json'
REQUIREMENTS_URL = 'https://raw.githubusercontent.com/Deutscher775/s.jesn.zip-tray-agent/refs/heads/main/requirements.txt'
SCRIPT_URL = 'https://raw.githubusercontent.com/Deutscher775/s.jesn.zip-tray-agent/refs/heads/main/JesnZIP-tray.py'
ICON_URL = 'https://raw.githubusercontent.com/Deutscher775/s.jesn.zip-tray-agent/refs/heads/main/ICON.ico'
SETTINGS_URL = 'https://raw.githubusercontent.com/Deutscher775/s.jesn.zip-tray-agent/refs/heads/main/tray_settings.json'


def info(msg: str) -> None:
    print(msg)


def download(url: str, dest: Path) -> None:
    info(f"Downloading {url} -> {dest.name}")
    try:
        urllib.request.urlretrieve(url, str(dest))
    except Exception as e:
        raise RuntimeError(f"Failed to download {url}: {e}")


def ensure_venv(py_exe: str = sys.executable) -> Path:
    """Create or reuse a virtual environment and return path to its python executable.

    Note: when this script is bundled as a one-file exe (PyInstaller), `sys.executable`
    refers to the frozen exe inside a temp folder and cannot be used as the base
    interpreter for a venv. In that case, we attempt to find a system Python
    (e.g. `py -3` or `python`) and use it to create the venv.
    """
    def find_system_python() -> str | None:
        # Prefer Python 3.9 specifically. Return the full command list to invoke it.
        candidates = [
            ['py', '-3.9'],
            ['py', '-3.10'],
            ['py', '-3'],
            ['python']
        ]
        for cmd in candidates:
            try:
                # ask for minor version
                test = cmd + ['-c', 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")']
                res = subprocess.run(test, capture_output=True, text=True)
                if res.returncode == 0:
                    out = (res.stdout or res.stderr).strip()
                    if out.startswith('3.9'):
                        return cmd
            except Exception:
                continue
        # If we didn't find 3.9, fall back to any Python 3 (but warn).
        for cmd in candidates:
            try:
                test = cmd + ['-c', 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")']
                res = subprocess.run(test, capture_output=True, text=True)
                if res.returncode == 0:
                    out = (res.stdout or res.stderr).strip()
                    if out.startswith('3'):
                        info(f"Found Python {out} but it's not 3.9. Build may not match target runtime.")
                        return cmd
            except Exception:
                continue
        return None

    # Determine the desired python command (prefer 3.9)
    desired_cmd = None
    try:
        desired_cmd = find_system_python()
    except Exception:
        desired_cmd = None

    recreate = False
    if VENV_DIR.exists():
        # inspect existing venv's python version
        try:
            if os.name == 'nt':
                existing_python = VENV_DIR / 'Scripts' / 'python.exe'
            else:
                existing_python = VENV_DIR / 'bin' / 'python'
            if existing_python.exists():
                res = subprocess.run([str(existing_python), '-c', 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")'], capture_output=True, text=True)
                out = (res.stdout or res.stderr or '').strip()
                if not out.startswith('3.9'):
                    info(f"Existing venv python is {out}; recreating venv to target Python 3.9")
                    recreate = True
            else:
                recreate = True
        except Exception:
            recreate = True

    if recreate or not VENV_DIR.exists():
        info(f"Creating virtual environment at {VENV_DIR}")
        # If we have a desired system python command (preferably 3.9), use it to create the venv
        if desired_cmd:
            cmd = list(desired_cmd) + ['-m', 'venv', str(VENV_DIR)]
            run(cmd)
        else:
            # fallback to venv.EnvBuilder when no system python command found
            venv_builder = venv.EnvBuilder(with_pip=True)
            venv_builder.create(str(VENV_DIR))
    else:
        info(f"Using existing virtual environment at {VENV_DIR}")

    if os.name == 'nt':
        venv_python = VENV_DIR / 'Scripts' / 'python.exe'
    else:
        venv_python = VENV_DIR / 'bin' / 'python'
    if not venv_python.exists():
        raise FileNotFoundError(f"Virtual environment python not found at {venv_python}")
    return venv_python


def run(cmd, *, check=True, capture_output=False, env=None):
    info('> ' + ' '.join(map(str, cmd)))
    res = subprocess.run(cmd, check=False, capture_output=capture_output, env=env)
    if check and res.returncode != 0:
        stdout = res.stdout.decode(errors='ignore') if res.stdout else ''
        stderr = res.stderr.decode(errors='ignore') if res.stderr else ''
        raise subprocess.CalledProcessError(res.returncode, cmd, output=stdout, stderr=stderr)
    return res


def install_requirements(venv_python: Path) -> None:
    info('Upgrading pip/setuptools/wheel in venv')
    run([str(venv_python), '-m', 'pip', 'install', '--upgrade', 'pip', 'setuptools', 'wheel'])

    info('Installing requirements into venv')
    # Ensure requirements.txt is present; if not, download
    req_path = ROOT / 'requirements.txt'
    if not req_path.exists():
        download(REQUIREMENTS_URL, req_path)

    try:
        run([str(venv_python), '-m', 'pip', 'install', '-r', str(req_path)])
    except subprocess.CalledProcessError as e:
        info('pip install -r failed; attempting per-package installs and skipping optional "winrt" if necessary')
        with req_path.open('r', encoding='utf-8') as fh:
            pkgs = [line.strip() for line in fh if line.strip() and not line.strip().startswith('#')]
        for pkg in pkgs:
            try:
                info(f'Installing package: {pkg}')
                run([str(venv_python), '-m', 'pip', 'install', pkg])
            except subprocess.CalledProcessError:
                if pkg.lower().startswith('winrt') or pkg == 'winrt':
                    info('Optional package winrt failed to install; continuing without native notifications')
                else:
                    raise RuntimeError(f'Critical package failed to install: {pkg}')


def uninstall_blacklist(venv_python: Path) -> None:
    blacklist = ['typing', 'pathlib']
    for pkg in blacklist:
        try:
            res = subprocess.run([str(venv_python), '-m', 'pip', 'show', pkg], stdout=subprocess.DEVNULL)
            if res.returncode == 0:
                info(f'Found incompatible package "{pkg}" in venv; uninstalling...')
                run([str(venv_python), '-m', 'pip', 'uninstall', '-y', pkg], check=False)
        except Exception as e:
            info(f'Check/uninstall for package {pkg} raised an exception: {e}')


def run_pyinstaller(venv_python: Path, script: Path, icon: Path | None) -> Path:
    info('Running PyInstaller from venv')
    # Prefer installed pyinstaller.exe if present in Scripts, else use module
    if os.name == 'nt':
        pyinstaller_exe = VENV_DIR / 'Scripts' / 'pyinstaller.exe'
    else:
        raise NotImplementedError('This script currently only supports Windows builds.')

    if pyinstaller_exe.exists():
        cmd = [str(pyinstaller_exe), '--noconfirm', '--onefile', '--windowed']
        if icon and icon.exists():
            # Set exe icon and include the icon file as bundled data so runtime can load it
            cmd.append(f'--icon={str(icon)}')
            if os.name == 'nt':
                data_arg = f"{str(icon)};."  # Windows: SRC;DEST
            else:
                data_arg = f"{str(icon)}:."  # POSIX: SRC:DEST
            cmd.extend(['--add-data', data_arg])
        cmd.append(str(script))
        run(cmd)
    else:
        cmd = [str(venv_python), '-m', 'PyInstaller', '--noconfirm', '--onefile', '--windowed']
        if icon and icon.exists():
            cmd.extend(['--icon', str(icon)])
            if os.name == 'nt':
                data_arg = f"{str(icon)};."  # Windows: SRC;DEST
            else:
                data_arg = f"{str(icon)}:."  # POSIX: SRC:DEST
            cmd.extend(['--add-data', data_arg])
        cmd.append(str(script))
        run(cmd)

    dist_dir = ROOT / 'dist'
    if not dist_dir.exists():
        raise FileNotFoundError('dist directory not found; PyInstaller may have failed')
    if icon and icon.exists():
        try:
            dest_icon = dist_dir / icon.name
            shutil.copy2(icon, dest_icon)
            info(f'Copied icon {icon.name} to {dest_icon}')
        except Exception as e:
            info(f'Failed to copy icon to dist: {e}')
    exe = None
    for f in dist_dir.glob('*.exe'):
        exe = f
        break
    if not exe:
        raise FileNotFoundError('No executable found in dist; build may have failed')
    return exe


def create_zip(exe_path: Path) -> Path:
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    zip_name = ROOT / f'JesnZIP-tray-{timestamp}.zip'
    info(f'Creating zip {zip_name}')
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe_path, exe_path.name)
    return zip_name


def create_start_menu_shortcut(exe_path: Path) -> None:
    # Try to create a Start Menu shortcut via PowerShell WScript.Shell COM (works without pywin32)
    start_menu = Path(os.environ.get('APPDATA', '')) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs'
    shortcut_dir = start_menu / 'JesnZIP'
    shortcut_dir.mkdir(parents=True, exist_ok=True)
    link_path = shortcut_dir / 'JesnZIP-tray.lnk'
    ps_cmd = (
        "$wsh = New-Object -ComObject WScript.Shell;"
        f"$s = $wsh.CreateShortcut('{str(link_path)}');"
        f"$s.TargetPath = '{str(exe_path)}';"
        f"$s.WorkingDirectory = '{str(exe_path.parent)}';"
        f"$s.IconLocation = '{str(exe_path)}';"
        "$s.Description = 'JesnZIP tray agent'; $s.Save();"
    )
    info('Creating Start Menu shortcut (via PowerShell COM)')
    run(['powershell', '-NoProfile', '-Command', ps_cmd], check=False)
    info(f'Start Menu shortcut attempted at: {link_path}')


def main() -> None:
    os.chdir(str(ROOT))
    info(f'Running installer.py in: {ROOT}')

    # Ensure required files are present (download if missing)
    if not (ROOT / SCRIPT_NAME).exists():
        download(SCRIPT_URL, ROOT / SCRIPT_NAME)
    if not (ROOT / ICON_NAME).exists():
        try:
            download(ICON_URL, ROOT / ICON_NAME)
        except Exception:
            info('ICON.ico not available; continuing without icon')
    if not (ROOT / SETTINGS_NAME).exists():
        try:
            download(SETTINGS_URL, ROOT / SETTINGS_NAME)
        except Exception:
            info('tray_settings.json not available; continuing')

    # Ensure requirements.txt exists (downloaded by install_requirements if missing)
    venv_python = ensure_venv()
    install_requirements(venv_python)
    uninstall_blacklist(venv_python)

    exe = run_pyinstaller(venv_python, ROOT / SCRIPT_NAME, ROOT / ICON_NAME)
    zip_path = create_zip(exe)
    info(f'Built and zipped: {zip_path}')

    create_shortcut = input('Create Start Menu shortcut? (Y/n) [Default: Y] ').strip()
    if create_shortcut == '' or create_shortcut.lower().startswith('y'):
        create_start_menu_shortcut(exe)
    else:
        info('Skipping Start Menu shortcut creation as requested.')

    login_prompt = input('Do you want to login with a session key now? (y/N) [Default: N] ').strip()
    if login_prompt.lower().startswith('y'):
        info('Launching JesnZIP-tray for session key login...')
        run([str(exe), '--session-prompt'], check=False)
    else:
        info('Skipping session key login prompt as requested.')

    try:
        subprocess.Popen([str(exe)], cwd=str(exe.parent), close_fds=True)
    except Exception as e:
        info(f'Failed to launch JesnZIP-tray: {e}')
    
    info('Installer completed successfully :3')
    sys.exit(0)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('ERROR:', e)
        sys.exit(1)
