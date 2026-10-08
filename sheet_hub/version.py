from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import requests


APP_VERSION = "1.4.0"
UPDATE_REPO = "christiancagfr-alt/SheetDataHub"
RELEASES_URL = f"https://github.com/{UPDATE_REPO}/releases"
INSTALLER_URL_PREFIXES = (
    "https://github.com/",
    "https://objects.githubusercontent.com/",
    "https://release-assets.githubusercontent.com/",
    "https://github-releases.githubusercontent.com/",
)


def parse_version(text: str) -> tuple[int, ...]:
    cleaned = str(text or "").strip().lstrip("vV")
    parts: list[int] = []
    for item in re.split(r"[.\-+_]", cleaned):
        if item.isdigit():
            parts.append(int(item))
        elif parts:
            break
    return tuple(parts or (0,))


def is_newer(latest: str, current: str = APP_VERSION) -> bool:
    left, right = parse_version(latest), parse_version(current)
    width = max(len(left), len(right))
    return left + (0,) * (width - len(left)) > right + (0,) * (width - len(right))


def pick_windows_installer(assets: list[dict[str, Any]]) -> dict[str, Any] | None:
    ranked: list[tuple[int, dict[str, Any]]] = []
    for asset in assets:
        name = str(asset.get("name") or "").casefold()
        if not name.endswith(".exe") or "sheetdatahub" not in name or "setup" not in name:
            continue
        ranked.append((0 if re.search(r"-v\d", name) else 1, asset))
    if not ranked:
        return None
    ranked.sort(key=lambda item: item[0])
    return ranked[0][1]


def installer_url_allowed(url: str) -> bool:
    text = str(url or "").strip()
    if not text.startswith("https://"):
        return False
    return any(text.startswith(prefix) for prefix in INSTALLER_URL_PREFIXES)


def fetch_latest_release(timeout: int = 20) -> dict[str, Any]:
    response = requests.get(
        f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",
        timeout=timeout,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SheetDataHub",
        },
    )
    if response.status_code == 404:
        raise RuntimeError("还没有发布版本，稍后再试。")
    response.raise_for_status()
    data = response.json()
    tag = str(data.get("tag_name") or "").strip()
    version = tag.lstrip("vV") or APP_VERSION
    assets = list(data.get("assets") or [])
    installer = pick_windows_installer(assets)
    installer_url = str((installer or {}).get("browser_download_url") or "")
    if installer_url and not installer_url_allowed(installer_url):
        installer_url = ""
    return {
        "tag": tag or f"v{version}",
        "version": version,
        "url": str(data.get("html_url") or RELEASES_URL),
        "name": str(data.get("name") or tag or version),
        "installer_url": installer_url,
        "installer_size": int((installer or {}).get("size") or 0),
    }


def download_release_installer(
    info: dict[str, Any],
    data_dir: str | Path,
    timeout: int = 180,
) -> Path:
    url = str(info.get("installer_url") or "").strip()
    if not url:
        raise RuntimeError("该版本没有可用的 Windows 安装包，请稍后再试。")
    if not installer_url_allowed(url):
        raise RuntimeError("安装包地址不是 GitHub 官方下载链接，已拒绝。")
    version = re.sub(r"[^0-9A-Za-z._-]", "_", str(info.get("version") or "latest"))
    update_dir = Path(data_dir) / "updates"
    update_dir.mkdir(parents=True, exist_ok=True)
    destination = update_dir / f"SheetDataHub-Setup-v{version}.exe"
    partial = destination.with_suffix(".exe.part")
    expected_size = int(info.get("installer_size") or 0)
    try:
        with requests.get(
            url,
            stream=True,
            timeout=(20, timeout),
            headers={"User-Agent": "SheetDataHub"},
        ) as response:
            response.raise_for_status()
            final_url = str(getattr(response, "url", url) or url)
            if not installer_url_allowed(final_url):
                raise RuntimeError("安装包下载被重定向到非 GitHub 地址，已拒绝。")
            written = 0
            with partial.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        output.write(chunk)
                        written += len(chunk)
        if expected_size and written != expected_size:
            raise RuntimeError(f"安装包下载不完整：应为 {expected_size} 字节，实际 {written} 字节。")
        if written < 1024 * 1024:
            raise RuntimeError("下载到的安装包大小异常。")
        with partial.open("rb") as source:
            if source.read(2) != b"MZ":
                raise RuntimeError("下载内容不是有效的 Windows 安装程序。")
        partial.replace(destination)
        return destination
    except Exception:
        partial.unlink(missing_ok=True)
        raise
