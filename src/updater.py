"""
Mô-đun kiểm tra và hỗ trợ cập nhật tự động từ GitHub Releases.
Sử dụng thư viện tiêu chuẩn urllib, tích hợp cơ chế tự động xử lý chứng chỉ SSL trên macOS.
"""
from __future__ import annotations

import json
import re
import ssl
import threading
import urllib.request
import urllib.error
import webbrowser
from src.version import APP_VERSION, GITHUB_REPO

def parse_version_tuple(v_str: str) -> tuple[int, ...]:
    """Chuyển đổi chuỗi phiên bản dạng 'v1.2.15' hoặc '1.2.0' thành tuple (1, 2, 15) để so sánh số học."""
    if not v_str:
        return (0, 0, 0)
    # Lấy các nhóm chữ số
    numbers = re.findall(r'\d+', str(v_str))
    if not numbers:
        return (0, 0, 0)
    return tuple(int(x) for x in numbers)

def _open_url_with_ssl_fallback(req: urllib.request.Request, timeout: int = 5):
    """
    Thực hiện gửi yêu cầu HTTPS với cơ chế tự động xử lý lỗi SSL chứng chỉ:
    1. Ưu tiên kiểm tra chứng chỉ bằng CA bundle chuẩn từ gói certifi (nếu có).
    2. Thử SSL Context mặc định của hệ thống.
    3. Nếu gặp lỗi thiếu chứng chỉ cục bộ (rất phổ biến trên macOS do chưa chạy Install Certificates
       hoặc trong file đóng gói PyInstaller), sẽ tự động fallback sang context unverified để đọc
       dữ liệu công khai từ GitHub API mà không bị lỗi CERTIFICATE_VERIFY_FAILED.
    """
    ssl_context = None

    # Thử nạp chứng chỉ CA qua certifi nếu môi trường có cài đặt
    try:
        import certifi
        ssl_context = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            ssl_context = ssl.create_default_context()
        except Exception:
            pass

    try:
        if ssl_context:
            return urllib.request.urlopen(req, timeout=timeout, context=ssl_context)
        return urllib.request.urlopen(req, timeout=timeout)
    except Exception as exc:
        err_str = str(exc)
        # Bắt lỗi CERTIFICATE_VERIFY_FAILED đặc trưng trên macOS hoặc môi trường PyInstaller
        if "CERTIFICATE_VERIFY_FAILED" in err_str or "certificate verify failed" in err_str:
            unverified_ctx = getattr(ssl, "_create_unverified_context", None)
            if unverified_ctx:
                return urllib.request.urlopen(req, timeout=timeout, context=unverified_ctx())
        raise

def check_for_updates(
    current_version: str = APP_VERSION,
    repo: str = GITHUB_REPO,
    timeout: int = 5
) -> dict:
    """
    Kiểm tra phiên bản mới nhất từ GitHub Releases API:
    Trả về dict:
    {
        "success": bool,
        "has_update": bool,
        "current_version": str,
        "latest_version": str,
        "release_name": str,
        "release_url": str,
        "release_notes": str,
        "assets": list[dict],
        "error": str | None
    }
    """
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    result = {
        "success": False,
        "has_update": False,
        "current_version": current_version,
        "latest_version": current_version,
        "release_name": "",
        "release_url": f"https://github.com/{repo}/releases/latest",
        "release_notes": "",
        "assets": [],
        "error": None
    }

    try:
        req = urllib.request.Request(
            api_url,
            headers={
                "User-Agent": "SohoaTHCS-Updater",
                "Accept": "application/vnd.github.v3+json"
            }
        )
        with _open_url_with_ssl_fallback(req, timeout=timeout) as resp:
            if resp.status != 200:
                result["error"] = f"HTTP Error {resp.status}"
                return result
            data = json.loads(resp.read().decode('utf-8'))

        latest_tag = data.get("tag_name", "")
        result["latest_version"] = latest_tag
        result["release_name"] = data.get("name", latest_tag)
        result["release_url"] = data.get("html_url", result["release_url"])
        result["release_notes"] = data.get("body", "")

        assets_list = []
        for asset in data.get("assets", []):
            assets_list.append({
                "name": asset.get("name", ""),
                "size": asset.get("size", 0),
                "download_url": asset.get("browser_download_url", "")
            })
        result["assets"] = assets_list

        curr_t = parse_version_tuple(current_version)
        latest_t = parse_version_tuple(latest_tag)

        # Nếu phiên bản mới trên GitHub lớn hơn phiên bản hiện tại
        if latest_t > curr_t:
            result["has_update"] = True

        result["success"] = True
    except Exception as e:
        result["error"] = str(e)

    return result

def check_for_updates_async(
    callback,
    current_version: str = APP_VERSION,
    repo: str = GITHUB_REPO,
    timeout: int = 5
):
    """Chạy kiểm tra cập nhật trên luồng ngầm (daemon thread) để không làm đơ giao diện GUI."""
    def _worker():
        res = check_for_updates(current_version, repo, timeout)
        if callback:
            callback(res)

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    return t

def open_release_page(url: str = f"https://github.com/{GITHUB_REPO}/releases/latest"):
    """Mở trang tải bản cập nhật trên trình duyệt web mặc định của hệ điều hành."""
    try:
        webbrowser.open(url)
    except Exception:
        pass
