from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests


APP_VERSION = "1.4.2"
UPDATE_REPO = "christiancagfr-alt/SheetDataHub"
RELEASES_URL = f"https://github.com/{UPDATE_REPO}/releases"
INSTALLER_DOWNLOAD_PREFIX = f"https://github.com/{UPDATE_REPO}/releases/download/"
INSTALLER_CDN_HOSTS = {
    "objects.githubusercontent.com",
    "release-assets.githubusercontent.com",
    "github-releases.githubusercontent.com",
}
MIN_INSTALLER_BYTES = 1024 * 1024
MAX_INSTALLER_BYTES = 400 * 1024 * 1024


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


def _asset_filename(asset: dict[str, Any]) -> str:
    name = str(asset.get("name") or "").strip()
    if name:
        return name
    path = urlparse(str(asset.get("browser_download_url") or "")).path
    return Path(path).name


def _rank_asset(asset: dict[str, Any], prefer_tokens: tuple[str, ...] = ()) -> tuple[int, dict[str, Any]]:
    name = _asset_filename(asset).casefold()
    score = 0 if re.search(r"-v\d", name) else 10
    for index, token in enumerate(prefer_tokens):
        if token and token not in name:
            score += 20 + index
    return (score, asset)


def pick_named_asset(
    assets: list[dict[str, Any]],
    *,
    suffixes: tuple[str, ...],
    required_tokens: tuple[str, ...],
    prefer_tokens: tuple[str, ...] = (),
) -> dict[str, Any] | None:
    ranked: list[tuple[int, dict[str, Any]]] = []
    for asset in assets:
        name = _asset_filename(asset).casefold()
        if not any(name.endswith(suffix) for suffix in suffixes):
            continue
        if any(token not in name for token in required_tokens):
            continue
        ranked.append(_rank_asset(asset, prefer_tokens))
    if not ranked:
        return None
    ranked.sort(key=lambda item: item[0])
    return ranked[0][1]


def pick_windows_installer(assets: list[dict[str, Any]]) -> dict[str, Any] | None:
    return pick_named_asset(
        assets,
        suffixes=(".exe",),
        required_tokens=("sheetdatahub", "setup"),
    )


def pick_macos_installer(assets: list[dict[str, Any]]) -> dict[str, Any] | None:
    machine = _macos_machine().casefold()
    prefer = ("arm64",) if ("arm" in machine or "aarch" in machine) else ("x86_64", "intel", "amd64")
    dmg = pick_named_asset(
        assets,
        suffixes=(".dmg",),
        required_tokens=("sheetdatahub",),
        prefer_tokens=prefer,
    )
    if dmg:
        return dmg
    return pick_named_asset(
        assets,
        suffixes=(".zip",),
        required_tokens=("sheetdatahub", "macos"),
        prefer_tokens=prefer,
    )


def _macos_machine() -> str:
    try:
        import platform
        return str(platform.machine() or "")
    except Exception:
        return ""


def pick_current_installer(assets: list[dict[str, Any]], platform_name: str | None = None) -> dict[str, Any] | None:
    platform_name = (platform_name or sys.platform).casefold()
    if platform_name.startswith("win"):
        return pick_windows_installer(assets)
    if platform_name == "darwin":
        return pick_macos_installer(assets)
    return None


def installer_kind(name: str, url: str = "") -> str:
    filename = str(name or "").strip() or Path(urlparse(url).path).name
    folded = filename.casefold()
    if folded.endswith(".exe"):
        return "exe"
    if folded.endswith(".dmg"):
        return "dmg"
    if folded.endswith(".zip"):
        return "zip"
    return ""


def _https_host(url: str) -> str:
    parsed = urlparse(str(url or "").strip())
    if parsed.scheme != "https" or parsed.username or parsed.password:
        return ""
    if parsed.port not in (None, 443):
        return ""
    return (parsed.hostname or "").casefold()


def installer_url_allowed(url: str) -> bool:
    text = str(url or "").strip()
    if not text.startswith(INSTALLER_DOWNLOAD_PREFIX):
        return False
    return _https_host(text) == "github.com"


def installer_final_url_allowed(url: str) -> bool:
    if installer_url_allowed(url):
        return True
    host = _https_host(url)
    return bool(host) and host in INSTALLER_CDN_HOSTS


def fetch_latest_release(timeout: int = 20, platform_name: str | None = None) -> dict[str, Any]:
    api_url = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
    response = requests.get(
        api_url,
        timeout=timeout,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SheetDataHub",
        },
    )
    final_api_url = str(getattr(response, "url", "") or api_url)
    if _https_host(final_api_url) != "api.github.com":
        raise RuntimeError("更新接口被重定向到非 GitHub 地址，已拒绝。")
    if response.status_code == 404:
        raise RuntimeError("还没有发布版本，稍后再试。")
    response.raise_for_status()
    data = response.json()
    if not isinstance(data, dict):
        raise RuntimeError("更新接口返回内容无效。")
    tag = str(data.get("tag_name") or "").strip()
    version = tag.lstrip("vV") or APP_VERSION
    assets = list(data.get("assets") or [])
    installer = pick_current_installer(assets, platform_name)
    installer_name = _asset_filename(installer or {})
    installer_url = str((installer or {}).get("browser_download_url") or "")
    kind = installer_kind(installer_name, installer_url)
    if installer_url and not installer_url_allowed(installer_url):
        installer_url = ""
        kind = ""
    return {
        "tag": tag or f"v{version}",
        "version": version,
        "url": str(data.get("html_url") or RELEASES_URL),
        "name": str(data.get("name") or tag or version),
        "installer_url": installer_url,
        "installer_name": installer_name,
        "installer_kind": kind,
        "installer_size": int((installer or {}).get("size") or 0),
    }


def _validate_downloaded_installer(path: Path, kind: str) -> None:
    if kind == "exe":
        with path.open("rb") as source:
            if source.read(2) != b"MZ":
                raise RuntimeError("下载内容不是有效的 Windows 安装程序。")
        return
    if kind == "zip":
        with path.open("rb") as source:
            if source.read(2) != b"PK":
                raise RuntimeError("下载内容不是有效的压缩包。")
        return
    if kind == "dmg":
        size = path.stat().st_size
        with path.open("rb") as source:
            source.seek(max(0, size - 512))
            tail = source.read(512)
        if b"koly" not in tail:
            raise RuntimeError("下载内容不是有效的 macOS 安装盘。")
        return
    raise RuntimeError("该版本没有当前系统可用的安装包。")


def download_release_installer(
    info: dict[str, Any],
    data_dir: str | Path,
    timeout: int = 180,
) -> Path:
    url = str(info.get("installer_url") or "").strip()
    if not url:
        raise RuntimeError("该版本没有当前系统可用的安装包，请稍后再试。")
    if not installer_url_allowed(url):
        raise RuntimeError("安装包地址不是本仓库的 GitHub 下载链接，已拒绝。")
    version = re.sub(r"[^0-9A-Za-z._-]", "_", str(info.get("version") or "latest"))
    kind = installer_kind(str(info.get("installer_name") or ""), url)
    if kind not in {"exe", "dmg", "zip"}:
        raise RuntimeError("该版本没有当前系统可用的安装包，请稍后再试。")
    update_dir = (Path(data_dir) / "updates").resolve()
    update_dir.mkdir(parents=True, exist_ok=True)
    destination = (update_dir / f"SheetDataHub-update-v{version}.{kind}").resolve()
    if destination.parent != update_dir:
        raise RuntimeError("安装包路径不在本机更新目录，已拒绝。")
    partial = destination.with_suffix(destination.suffix + ".part")
    expected_size = int(info.get("installer_size") or 0)
    if expected_size and expected_size > MAX_INSTALLER_BYTES:
        raise RuntimeError("安装包过大，已拒绝下载。")
    try:
        with requests.get(
            url,
            stream=True,
            timeout=(20, timeout),
            headers={"User-Agent": "SheetDataHub"},
        ) as response:
            response.raise_for_status()
            final_url = str(getattr(response, "url", url) or url)
            if not installer_final_url_allowed(final_url):
                raise RuntimeError("安装包下载被重定向到非 GitHub 地址，已拒绝。")
            written = 0
            with partial.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        written += len(chunk)
                        if written > MAX_INSTALLER_BYTES:
                            raise RuntimeError("安装包过大，已中止下载。")
                        output.write(chunk)
        if expected_size and written != expected_size:
            raise RuntimeError(f"安装包下载不完整：应为 {expected_size} 字节，实际 {written} 字节。")
        if written < MIN_INSTALLER_BYTES:
            raise RuntimeError("下载到的安装包大小异常。")
        _validate_downloaded_installer(partial, kind)
        partial.replace(destination)
        return destination
    except Exception:
        partial.unlink(missing_ok=True)
        raise
