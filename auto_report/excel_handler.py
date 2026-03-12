import subprocess
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client
import win32process


def get_running_excel_apps() -> list[Any]:
    apps: dict[int, Any] = {}
    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        bind_ctx = pythoncom.CreateBindCtx(0)
    except Exception:
        enum = None
        bind_ctx = None
        rot = None

    if enum is not None:
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break

            moniker = monikers[0]
            try:
                display_name = moniker.GetDisplayName(bind_ctx, None)
            except Exception:
                continue

            lower_name = display_name.lower()
            if "excel" not in lower_name and ".xls" not in lower_name:
                continue

            try:
                obj = rot.GetObject(moniker)
                app = win32com.client.Dispatch(obj).Application
                apps[int(app.Hwnd)] = app
            except Exception:
                continue

    if not apps:
        try:
            app = win32com.client.GetObject(Class="Excel.Application")
            apps[int(app.Hwnd)] = app
        except Exception:
            pass

    return list(apps.values())


def get_open_workbooks(apps: list[Any]) -> list[tuple[int, Any]]:
    open_workbooks: list[tuple[int, Any]] = []
    for app in apps:
        try:
            workbooks = app.Workbooks
            if workbooks.Count == 0:
                continue

            app_hwnd = int(app.Hwnd)
            for index in range(1, workbooks.Count + 1):
                wb = workbooks.Item(index)
                open_workbooks.append((app_hwnd, wb))
        except Exception:
            continue
    return open_workbooks


def kill_empty_excel_processes(apps: list[Any]) -> int:
    killed = 0
    keep_pids: set[int] = set()
    empty_pids: set[int] = set()

    for app in apps:
        try:
            hwnd = int(app.Hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            workbooks = app.Workbooks
            if workbooks.Count > 0:
                keep_pids.add(int(pid))
            else:
                empty_pids.add(int(pid))
        except Exception:
            continue

    for pid in empty_pids:
        if pid in keep_pids:
            continue
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                capture_output=True,
                text=True,
                check=False,
            )
            killed += 1
        except Exception:
            continue
    return killed


def save_all_open_workbooks(
    save_dir: Path,
    max_scan_retries: int = 2,
    scan_retry_sleep_seconds: float = 0.5,
    logger: Any | None = None,
) -> list[Path]:
    save_dir.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        apps = get_running_excel_apps()
        if not apps:
            if logger is not None:
                logger.warning("没有检测到正在运行的Excel进程")
            return []

        open_workbooks: list[tuple[int, Any]] = []
        for attempt in range(max_scan_retries + 1):
            open_workbooks = get_open_workbooks(apps)
            if open_workbooks:
                break

            killed = kill_empty_excel_processes(apps)
            if killed <= 0:
                break

            if attempt < max_scan_retries:
                time.sleep(scan_retry_sleep_seconds)
                apps = get_running_excel_apps()
                if not apps:
                    break

        if not open_workbooks:
            if logger is not None:
                logger.warning("检测到Excel进程，但没有打开的工作簿")
            return []

        results: list[Path] = []
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        saved_count = 0

        for app_hwnd, wb in open_workbooks:
            name_stem = Path(wb.Name).stem
            suffix = Path(wb.Name).suffix or ".xlsx"
            target = save_dir / f"{name_stem}_{timestamp}_{app_hwnd}_{saved_count + 1}{suffix}"
            wb.SaveCopyAs(str(target))
            results.append(target)
            saved_count += 1
            if logger is not None:
                logger.info("保存成功: %s", target)

        if logger is not None:
            logger.info("共保存 %s 个工作簿", saved_count)
        return results
    finally:
        pythoncom.CoUninitialize()
