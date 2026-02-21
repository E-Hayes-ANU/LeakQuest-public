"""Auto-update module for LeakQuest.

Checks GitHub Releases for newer versions and handles downloading updates.
"""

import os
import sys
import subprocess
import tempfile

import requests
from rich.progress import Progress, BarColumn, DownloadColumn, TransferSpeedColumn

REPO = "E-Hayes-ANU/LeakQuest-public"
API_URL = f"https://api.github.com/repos/{REPO}/releases"


def _get_platform_suffix():
    """Return the release tag suffix for the current platform."""
    if sys.platform == "win32":
        return "-windows"
    elif sys.platform == "darwin":
        return "-mac"
    return ""


def _parse_version(tag):
    """Parse version tuple from a release tag like 'v2.2-windows'."""
    # Strip leading 'v' and trailing platform suffix
    ver = tag.lstrip("v")
    for suffix in ("-windows", "-mac"):
        if ver.endswith(suffix):
            ver = ver[: -len(suffix)]
            break
    try:
        return tuple(int(x) for x in ver.split("."))
    except (ValueError, AttributeError):
        return ()


def check_for_update(current_version):
    """Check GitHub for a newer release matching the current platform.

    Args:
        current_version: Current version string (e.g. "2.1")

    Returns:
        (new_version_str, download_url, asset_name) if an update is available,
        or None if up-to-date or on error.
    """
    platform_suffix = _get_platform_suffix()
    if not platform_suffix:
        return None

    current_tuple = _parse_version(f"v{current_version}{platform_suffix}")
    if not current_tuple:
        return None

    try:
        resp = requests.get(API_URL, timeout=10, headers={
            "Accept": "application/vnd.github.v3+json",
        })
        resp.raise_for_status()
        releases = resp.json()
    except Exception:
        return None

    # Find latest release for this platform
    best_version = current_tuple
    best_release = None

    for release in releases:
        tag = release.get("tag_name", "")
        if not tag.endswith(platform_suffix):
            continue
        ver = _parse_version(tag)
        if ver and ver > best_version:
            best_version = ver
            best_release = release

    if not best_release:
        return None

    # Find the download asset
    assets = best_release.get("assets", [])
    if not assets:
        return None

    asset = assets[0]
    new_version_str = ".".join(str(x) for x in best_version)
    return (new_version_str, asset["browser_download_url"], asset["name"])


def download_update(download_url, asset_name, console):
    """Download the update file with a progress bar.

    Windows: saves next to the running exe as LeakQuest_update.exe
    Mac: saves to ~/Downloads

    Args:
        download_url: URL to download from
        asset_name: Original filename of the asset
        console: rich Console instance for output

    Returns:
        Path to the downloaded file, or None on failure.
    """
    try:
        resp = requests.get(download_url, stream=True, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        console.print(f"[red]Download failed: {e}[/red]")
        return None

    total = int(resp.headers.get("content-length", 0))

    # Determine save location
    if sys.platform == "win32":
        if getattr(sys, "frozen", False):
            # Running as PyInstaller exe — save next to the exe
            dest_dir = os.path.dirname(sys.executable)
        else:
            dest_dir = os.path.dirname(os.path.abspath(__file__))
        dest_path = os.path.join(dest_dir, "LeakQuest_update.exe")
    else:
        dest_dir = os.path.expanduser("~/Downloads")
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, asset_name)

    try:
        with open(dest_path, "wb") as f:
            with Progress(
                BarColumn(),
                DownloadColumn(),
                TransferSpeedColumn(),
                console=console,
            ) as progress:
                task = progress.add_task("Downloading...", total=total or None)
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
                    progress.update(task, advance=len(chunk))
    except Exception as e:
        console.print(f"[red]Failed to save update: {e}[/red]")
        return None

    return dest_path


def apply_windows_update(dest_path, console):
    """Spawn a batch script to replace the running exe after exit.

    The batch script waits 2 seconds, deletes the old exe, renames the
    update exe, then deletes itself. Runs in a hidden window.
    """
    if not getattr(sys, "frozen", False):
        console.print(
            f"[yellow]Update saved to:[/yellow] {dest_path}\n"
            "[yellow]Replace the current script manually to complete the update.[/yellow]"
        )
        return

    exe_path = sys.executable
    exe_dir = os.path.dirname(exe_path)
    exe_name = os.path.basename(exe_path)
    update_name = os.path.basename(dest_path)

    bat_path = os.path.join(exe_dir, "_leakquest_update.bat")
    bat_content = (
        "@echo off\n"
        "timeout /t 2 /nobreak >nul\n"
        f'del "{exe_name}"\n'
        f'rename "{update_name}" "{exe_name}"\n'
        f'del "%~f0"\n'
    )

    with open(bat_path, "w") as f:
        f.write(bat_content)

    # Launch hidden — CREATE_NO_WINDOW = 0x08000000
    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        cwd=exe_dir,
        creationflags=CREATE_NO_WINDOW,
        close_fds=True,
    )

    console.print(
        "\n[bold green]Update downloaded successfully![/bold green]\n"
        "LeakQuest will now close. Re-open it to use the new version."
    )
