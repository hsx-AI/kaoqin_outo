import logging
import subprocess
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client
import win32process

logger = logging.getLogger("auto_report")

XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


def kill_all_excel_processes() -> int:
    """强制终止所有 EXCEL.EXE 进程，返回成功终止的数量"""
    killed = 0
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
            capture_output=True, text=True, check=False,
        )
        if result.returncode == 0:
            killed = max(result.stdout.upper().count("SUCCESS"), 1)
    except Exception:
        pass
    return killed


def _optimize_excel_for_save(app) -> dict:
    """保存前关闭 Excel 公式重算/屏幕刷新/事件/弹窗，大幅加速 SaveCopyAs"""
    original: dict[str, Any] = {}
    for attr, fast_val in [
        ("Calculation", XL_CALCULATION_MANUAL),
        ("ScreenUpdating", False),
        ("DisplayAlerts", False),
        ("EnableEvents", False),
    ]:
        try:
            original[attr] = getattr(app, attr)
            setattr(app, attr, fast_val)
        except Exception:
            pass
    return original


def _restore_excel_settings(app, original: dict) -> None:
    """SaveCopyAs 完成后恢复 Excel 原始设置"""
    for attr, val in original.items():
        try:
            setattr(app, attr, val)
        except Exception:
            pass


def get_running_excel_apps() -> list[Any]:
    """
    多策略检测正在运行的 Excel 实例：
    1. ROT (Running Object Table) 枚举
    2. GetActiveObject (获取最近激活的实例)
    3. GetObject with Class (OLE 类查找)
    """
    apps: dict[int, Any] = {}

    # 策略 1：ROT 枚举
    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        bind_ctx = pythoncom.CreateBindCtx(0)

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
    except Exception:
        pass

    # 策略 2：GetActiveObject
    if not apps:
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            apps[int(app.Hwnd)] = app
        except Exception:
            pass

    # 策略 3：GetObject with Class
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
                capture_output=True, text=True, check=False,
            )
            killed += 1
        except Exception:
            continue
    return killed


def save_all_open_workbooks(
    save_dir: Path,
    max_scan_retries: int = 5,
    scan_retry_sleep_seconds: float = 2.0,
    logger: Any | None = None,
) -> list[Path]:
    save_dir.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        open_workbooks: list[tuple[int, Any]] = []

        for attempt in range(max_scan_retries + 1):
            try:
                pythoncom.CoFreeUnusedLibraries()
            except Exception:
                pass

            apps = get_running_excel_apps()

            if not apps:
                if logger is not None:
                    logger.info("第 %d/%d 次扫描: 未检测到Excel进程",
                                attempt + 1, max_scan_retries + 1)
                if attempt < max_scan_retries:
                    time.sleep(scan_retry_sleep_seconds)
                continue

            if logger is not None:
                logger.info("第 %d/%d 次扫描: 检测到 %d 个Excel实例",
                            attempt + 1, max_scan_retries + 1, len(apps))

            open_workbooks = get_open_workbooks(apps)
            if open_workbooks:
                break

            killed = kill_empty_excel_processes(apps)
            if logger is not None and killed > 0:
                logger.info("清理了 %d 个空Excel进程", killed)

            if attempt < max_scan_retries:
                time.sleep(scan_retry_sleep_seconds)

        if not open_workbooks:
            if logger is not None:
                logger.warning("所有 %d 次扫描完成，未找到打开的工作簿",
                               max_scan_retries + 1)
            return []

        results: list[Path] = []
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        saved_count = 0

        for app_hwnd, wb in open_workbooks:
            name_stem = Path(wb.Name).stem
            suffix = Path(wb.Name).suffix or ".xlsx"
            target = save_dir / f"{name_stem}_{timestamp}_{app_hwnd}_{saved_count + 1}{suffix}"

            app = wb.Application
            original_settings = _optimize_excel_for_save(app)
            try:
                if logger is not None:
                    logger.info("正在保存 %s (已优化Excel设置)...", wb.Name)
                t0 = time.time()
                wb.SaveCopyAs(str(target))
                elapsed = time.time() - t0
                if logger is not None:
                    logger.info("保存成功 (%.1f秒): %s", elapsed, target)
            except Exception as e:
                if logger is not None:
                    logger.warning("SaveCopyAs 失败: %s, 等待5秒后重试...", e)
                time.sleep(5)
                try:
                    wb.SaveCopyAs(str(target))
                    if logger is not None:
                        logger.info("重试保存成功: %s", target)
                except Exception as e2:
                    if logger is not None:
                        logger.error("重试保存也失败: %s, 跳过此工作簿", e2)
                    continue
            finally:
                _restore_excel_settings(app, original_settings)

            results.append(target)
            saved_count += 1

        if logger is not None:
            logger.info("共保存 %d 个工作簿", saved_count)
        return results
    finally:
        pythoncom.CoUninitialize()
