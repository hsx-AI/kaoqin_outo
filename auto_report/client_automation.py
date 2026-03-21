from dataclasses import dataclass
import ctypes
import logging
import subprocess
import time
from pathlib import Path

import win32com.client
import win32clipboard
import win32con
import win32gui
import win32process
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
import re

logger = logging.getLogger("auto_report")


@dataclass
class ClientAutomation:
    executable_path: str | None = None
    pid: int | None = None
    hwnd: int | None = None
    process: subprocess.Popen | None = None

    # ── 辅助方法 ──────────────────────────────────────────────────────────

    def _find_visible_window_for_pid(self, pid: int) -> int | None:
        found_hwnd: int | None = None

        def on_window(hwnd: int, _lparam: int) -> bool:
            nonlocal found_hwnd
            try:
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if int(window_pid) != int(pid):
                    return True
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if not title.strip():
                    return True
                found_hwnd = hwnd
                return False
            except Exception:
                return True

        win32gui.EnumWindows(on_window, 0)
        return found_hwnd

    def _force_foreground_window(self, hwnd: int) -> bool:
        """通过 AttachThreadInput 技巧可靠地将窗口置前，解决 SetForegroundWindow 静默失败"""
        try:
            foreground_hwnd = win32gui.GetForegroundWindow()
            if foreground_hwnd == hwnd:
                return True

            current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
            foreground_thread = ctypes.windll.user32.GetWindowThreadProcessId(
                foreground_hwnd, None
            )

            attached = False
            if current_thread != foreground_thread and foreground_thread != 0:
                attached = bool(
                    ctypes.windll.user32.AttachThreadInput(
                        foreground_thread, current_thread, True
                    )
                )

            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
            finally:
                if attached:
                    ctypes.windll.user32.AttachThreadInput(
                        foreground_thread, current_thread, False
                    )

            time.sleep(0.3)
            return win32gui.GetForegroundWindow() == hwnd
        except Exception:
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
                return True
            except Exception:
                return False

    def _is_process_responsive(self, hwnd: int | None = None) -> bool:
        """检查窗口/进程是否正在响应（非挂起）"""
        target = hwnd or self.hwnd
        if not target:
            return True
        try:
            hung = ctypes.windll.user32.IsHungAppWindow(int(target))
            return not bool(hung)
        except Exception:
            return True

    def _check_responsive_or_raise(self, context: str = ""):
        """如果进程未响应则抛出 RuntimeError，用于中断潜在的死循环"""
        if self.hwnd and not self._is_process_responsive(self.hwnd):
            msg = f"程序未响应: {context}" if context else "程序未响应"
            raise RuntimeError(msg)

    def _find_child_by_text(self, parent_hwnd: int, target_text: str, fuzzy: bool = False) -> int | None:
        target = target_text.strip()
        if not target:
            return None
        found: int | None = None

        def walk(hwnd: int) -> None:
            nonlocal found
            if found is not None:
                return
            try:
                if not win32gui.IsWindow(hwnd):
                    return
                text = win32gui.GetWindowText(hwnd).strip()
                if (text == target) or (fuzzy and target in text):
                    found = int(hwnd)
                    return
            except Exception:
                return

            child = win32gui.FindWindowEx(hwnd, 0, None, None)
            while child:
                walk(child)
                if found is not None:
                    return
                child = win32gui.FindWindowEx(hwnd, child, None, None)

        walk(parent_hwnd)
        return found

    def _click_export_by_query_neighbor(self, parent_hwnd: int) -> bool:
        export_hwnd = self._find_child_by_text(parent_hwnd, "结果导出")
        if not export_hwnd:
            export_hwnd = self._find_child_by_text(parent_hwnd, "结果导出", fuzzy=True)
        if export_hwnd:
            try:
                win32gui.PostMessage(export_hwnd, win32con.BM_CLICK, 0, 0)
                return True
            except Exception:
                pass

        query_hwnd = self._find_child_by_text(parent_hwnd, "查询")
        if not query_hwnd:
            query_hwnd = self._find_child_by_text(parent_hwnd, "查询", fuzzy=True)
        if not query_hwnd:
            return False

        try:
            current = win32gui.GetWindow(query_hwnd, win32con.GW_HWNDNEXT)
        except Exception:
            return False

        for _ in range(30):
            if not current:
                break
            try:
                if win32gui.IsWindowVisible(current):
                    win32gui.PostMessage(current, win32con.BM_CLICK, 0, 0)
                    return True
            except Exception:
                pass
            try:
                current = win32gui.GetWindow(current, win32con.GW_HWNDNEXT)
            except Exception:
                break

        return False

    def _focus_export_confirm_dialog(self, main_window_title: str, timeout_seconds: float = 8.0) -> bool:
        deadline = time.time() + float(timeout_seconds)
        while time.time() < deadline:
            self._check_responsive_or_raise("等待导出确认弹窗")
            try:
                windows = Desktop(backend="uia").windows(process=self.pid, visible_only=True)
            except Exception:
                time.sleep(0.2)
                continue

            preferred = None
            fallback = None
            for window in windows:
                try:
                    title = (window.window_text() or "").strip()
                except Exception:
                    continue
                if not title:
                    continue
                if main_window_title in title:
                    continue
                if fallback is None:
                    fallback = window
                if "确认" in title or "导出" in title:
                    preferred = window
                    break

            target = preferred or fallback
            if target is not None:
                try:
                    target.set_focus()
                    return True
                except Exception:
                    try:
                        target_hwnd = int(getattr(target, "handle", 0) or 0)
                        if target_hwnd > 0:
                            win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                            win32gui.BringWindowToTop(target_hwnd)
                            win32gui.SetForegroundWindow(target_hwnd)
                            return True
                    except Exception:
                        pass
            time.sleep(0.2)
        return False

    # ── 启动 / 关闭 ──────────────────────────────────────────────────────

    def launch(self, timeout_seconds: float = 20.0) -> int:
        if not self.executable_path:
            raise ValueError("executable_path is required")

        exe_path = Path(self.executable_path)
        if not exe_path.exists():
            raise FileNotFoundError(str(exe_path))

        self.process = subprocess.Popen([str(exe_path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.pid = int(self.process.pid)

        deadline = time.time() + float(timeout_seconds)
        while time.time() < deadline:
            hwnd = self._find_visible_window_for_pid(self.pid)
            if hwnd is not None:
                self.hwnd = int(hwnd)
                return self.hwnd
            time.sleep(0.2)

        return 0

    def close_process_by_name(self, process_name: str) -> bool:
        closed = False

        if self.pid:
            try:
                result = subprocess.run(
                    ["taskkill", "/F", "/PID", str(int(self.pid)), "/T"],
                    capture_output=True, text=True, check=False,
                )
                if result.returncode == 0:
                    return True
            except Exception:
                pass

        names: list[str] = []
        base_name = process_name.strip()
        if base_name:
            names.append(base_name)
            if not base_name.lower().endswith(".exe"):
                names.append(f"{base_name}.exe")

        if self.executable_path:
            exe_name = Path(self.executable_path).name.strip()
            if exe_name:
                names.append(exe_name)

        deduped_names: list[str] = []
        seen: set[str] = set()
        for name in names:
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)
            deduped_names.append(name)

        for name in deduped_names:
            try:
                result = subprocess.run(
                    ["taskkill", "/F", "/IM", name, "/T"],
                    capture_output=True, text=True, check=False,
                )
                if result.returncode == 0:
                    closed = True
            except Exception:
                continue
        return closed

    def close_image_process(self, process_name: str) -> bool:
        base_name = process_name.strip()
        if not base_name:
            return False

        names = [base_name]
        if not base_name.lower().endswith(".exe"):
            names.append(f"{base_name}.exe")

        for name in names:
            try:
                result = subprocess.run(
                    ["taskkill", "/F", "/IM", name, "/T"],
                    capture_output=True, text=True, check=False,
                )
                if result.returncode == 0:
                    return True
            except Exception:
                continue
        return False

    def _set_clipboard_text(self, text: str) -> None:
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            raise

    # ── 登录（带重试 + 焦点验证）────────────────────────────────────────

    def login(self, username: str, password: str, timeout_seconds: float = 10.0,
             max_input_retries: int = 3) -> bool:
        """
        查找标题包含[登录]的窗口并输入凭据。
        使用 _force_foreground_window 确保焦点，失败自动重试。
        """
        if not self.pid:
            raise RuntimeError("Process not launched")

        login_hwnd: int | None = None
        deadline = time.time() + float(timeout_seconds)

        def on_window(hwnd: int, _lparam: int) -> bool:
            nonlocal login_hwnd
            try:
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if int(window_pid) != int(self.pid):
                    return True
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if "登录" in title or "Login" in title:
                    login_hwnd = hwnd
                    return False
                return True
            except Exception:
                return True

        while time.time() < deadline:
            win32gui.EnumWindows(on_window, 0)
            if login_hwnd:
                break
            time.sleep(0.5)

        if not login_hwnd:
            if self.hwnd and win32gui.IsWindowVisible(self.hwnd):
                login_hwnd = self.hwnd
            else:
                return False

        time.sleep(5.0)

        for attempt in range(max_input_retries):
            try:
                logger.info("登录输入第 %d/%d 次尝试", attempt + 1, max_input_retries)

                if not self._force_foreground_window(login_hwnd):
                    logger.warning("无法将登录窗口置前，等待后重试")
                    time.sleep(1.0)
                    continue

                shell = win32com.client.Dispatch("WScript.Shell")

                shell.SendKeys("^a")
                time.sleep(0.3)
                shell.SendKeys("{DELETE}")
                time.sleep(0.3)

                self._set_clipboard_text(username)
                time.sleep(0.3)

                # 粘贴前再次验证焦点
                if win32gui.GetForegroundWindow() != login_hwnd:
                    logger.warning("粘贴前焦点丢失，重新激活窗口")
                    if not self._force_foreground_window(login_hwnd):
                        time.sleep(1.0)
                        continue

                shell.SendKeys("^v")
                time.sleep(0.5)

                # 验证剪贴板未被其他程序篡改
                try:
                    win32clipboard.OpenClipboard()
                    clip_text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    if clip_text != username:
                        logger.warning("剪贴板内容被篡改，重试")
                        time.sleep(0.5)
                        continue
                except Exception:
                    try:
                        win32clipboard.CloseClipboard()
                    except Exception:
                        pass

                shell.SendKeys("{TAB}")
                time.sleep(0.3)

                shell.SendKeys(password)
                time.sleep(0.3)

                shell.SendKeys("{ENTER}")
                time.sleep(1.0)

                return True
            except Exception as e:
                logger.warning("登录输入第 %d 次失败: %s", attempt + 1, e)
                if attempt < max_input_retries - 1:
                    time.sleep(1.0)
                continue

        return False

    # ── 调试 ──────────────────────────────────────────────────────────────

    def _dump_debug_info(self, backend: str, filename: str):
        try:
            app = Application(backend=backend).connect(process=self.pid)
            windows = app.windows(visible_only=True)
            with open(filename, "w", encoding="utf-8") as f:
                import sys
                original_stdout = sys.stdout
                sys.stdout = f
                try:
                    for w in windows:
                        print(f"--- Window: {w.window_text()} ---")
                        try:
                            w.print_control_identifiers()
                        except Exception as e:
                            print(f"Error printing identifiers: {e}")
                finally:
                    sys.stdout = original_stdout
        except Exception as e:
            with open(filename, "a", encoding="utf-8") as f:
                f.write(f"Error dumping UI: {e}\n")

    # ── 原始数据查询（带响应性检测）──────────────────────────────────────

    def query_original_data(self, timeout_seconds: float = 10.0) -> bool:
        """查找并点击[原始数据查询]按钮，每轮循环检测程序是否挂起"""
        if not self.pid:
            raise RuntimeError("Process not launched")

        def try_click_with_backend(backend: str) -> bool:
            try:
                app = Application(backend=backend).connect(process=self.pid)
                windows = app.windows(title_re=".*先达考勤管理系统.*")

                for window in windows:
                    try:
                        window.set_focus()
                    except Exception:
                        pass

                    try:
                        btns = window.descendants(title="原始数据查询", control_type="Button")
                        for btn in btns:
                            try:
                                btn.click_input()
                                return True
                            except Exception:
                                pass

                        btns = window.descendants(title_re=".*原始数据查询.*", control_type="Button")
                        for btn in btns:
                            try:
                                btn.click_input()
                                return True
                            except Exception:
                                pass
                    except Exception:
                        pass

                    try:
                        menu_item = window.descendants(title="数据查询", control_type="MenuItem")
                        if menu_item:
                            menu_item[0].click_input()
                            time.sleep(0.5)
                            sub_item = window.descendants(title="原始数据查询", control_type="MenuItem")
                            if sub_item:
                                sub_item[0].click_input()
                                return True
                    except Exception:
                        pass

                return False
            except Exception:
                return False

        deadline = time.time() + float(timeout_seconds)
        while time.time() < deadline:
            self._check_responsive_or_raise("查找原始数据查询按钮")
            if try_click_with_backend("uia"):
                return True
            time.sleep(1.0)

        logger.warning("未找到按钮，正在导出UI结构到 ui_dump_uia.txt ...")
        self._dump_debug_info("uia", "ui_dump_uia.txt")
        return False

    # ── 查询导出（带全局超时 + 响应性检测）────────────────────────────────

    def perform_query_and_export(self, timeout_seconds: float = 120.0) -> bool:
        """
        执行查询并导出流程。全局超时保护，每步检测进程是否挂起。
        超时或进程挂起时抛出 RuntimeError 而不是静默死循环。
        """
        deadline = time.time() + float(timeout_seconds)

        def check_deadline(step: str):
            if time.time() > deadline:
                raise RuntimeError(f"操作超时 ({timeout_seconds}s): {step}")

        try:
            self._check_responsive_or_raise("连接应用")
            check_deadline("连接应用")
            app = Application(backend="uia").connect(process=self.pid)

            # 1. 点击[查询]按钮
            logger.info("正在查找主界面[查询]按钮...")
            main_window = app.window(title_re=".*先达考勤管理系统.*")
            main_window.set_focus()

            query_btn = main_window.descendants(title="查询", control_type="Button")
            if not query_btn:
                logger.error("未找到主界面[查询]按钮")
                return False

            query_btn[0].click_input()
            logger.info("点击主界面[查询]按钮")
            time.sleep(5)

            # 2. 在弹出的查询窗口中点击[确定]
            logger.info("正在等待[查询]窗口...")
            query_dialog = None
            for i in range(10):
                check_deadline("等待查询窗口")
                self._check_responsive_or_raise("等待查询窗口")
                try:
                    windows = Desktop(backend="uia").windows(process=self.pid, visible_only=True)
                    if i == 0:
                        logger.info("当前可见窗口: %s", [w.window_text() for w in windows])

                    for w in windows:
                        if "查询" in w.window_text():
                            query_dialog = w
                            break
                    if query_dialog:
                        break
                except Exception as e:
                    logger.warning("查找窗口出错: %s", e)
                time.sleep(1.0)

            if query_dialog:
                try:
                    query_dialog.set_focus()
                    confirm_btn = query_dialog.descendants(title="确定", control_type="Button")
                    if confirm_btn:
                        confirm_btn[0].click_input()
                        logger.info("点击查询窗口[确定]按钮")
                    else:
                        logger.info("未找到[确定]按钮，尝试按回车键...")
                        query_dialog.type_keys("{ENTER}")
                except Exception as e:
                    logger.warning("操作查询窗口失败: %s，尝试按回车键...", e)
                    try:
                        query_dialog.type_keys("{ENTER}")
                    except Exception:
                        pass
            else:
                logger.warning("未找到查询窗口(超时)，尝试在主窗口按回车键...")
                try:
                    main_window.type_keys("{ENTER}")
                except Exception:
                    pass

            # 3. 等待数据加载（分段等待，每段检查响应性）
            logger.info("等待数据加载...")
            for _ in range(5):
                check_deadline("等待数据加载")
                self._check_responsive_or_raise("等待数据加载")
                time.sleep(1.0)

            # 4. 点击[结果导出]
            logger.info("正在查找[结果导出]按钮...")
            check_deadline("查找结果导出")
            self._check_responsive_or_raise("查找结果导出")

            main_wrapper = None
            try:
                main_window = app.window(title_re=".*先达考勤.*")
                if not main_window.exists(timeout=5):
                    logger.info("未找到主窗口，尝试重新连接应用...")
                    app = Application(backend="uia").connect(process=self.pid)
                    main_window = app.window(title_re=".*先达考勤.*")

                main_wrapper = main_window.wrapper_object()
                wrapper_hwnd = int(getattr(main_wrapper, "handle", 0) or 0)
                if wrapper_hwnd > 0:
                    self.hwnd = wrapper_hwnd
                elif self.pid:
                    fallback_hwnd = self._find_visible_window_for_pid(int(self.pid))
                    if fallback_hwnd:
                        self.hwnd = int(fallback_hwnd)
                try:
                    main_wrapper.set_focus()
                    logger.info("已定位主窗口: %s", main_wrapper.window_text())
                except Exception as e:
                    logger.warning("主窗口set_focus失败(忽略): %s", e)
            except Exception as e:
                logger.error("重新查找主窗口失败: %s", e)
                return False

            parent_hwnd = int(self.hwnd) if self.hwnd else 0
            if parent_hwnd <= 0 and self.pid:
                temp_hwnd = self._find_visible_window_for_pid(int(self.pid))
                if temp_hwnd:
                    parent_hwnd = int(temp_hwnd)
                    self.hwnd = parent_hwnd

            if parent_hwnd <= 0:
                logger.error("未找到主窗口句柄，无法执行Win32相邻按钮点击")
                return False

            if not self._click_export_by_query_neighbor(parent_hwnd):
                logger.error("Win32相邻按钮点击未命中")
                return False

            logger.info("已执行Win32相邻按钮点击")

            # 5. 导出确认弹窗
            logger.info("等待确认弹窗出现...")
            time.sleep(3)
            check_deadline("确认导出弹窗")

            logger.info("正在确认导出(发送回车)...")
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys("{ENTER}")
            except Exception:
                try:
                    main_window.type_keys("{ENTER}")
                except Exception:
                    pass

            # 6. 等待导出完成（分段等待，每段检查响应性和超时）
            logger.info("等待导出完成...")
            for _ in range(15):
                check_deadline("等待导出完成")
                self._check_responsive_or_raise("等待导出完成")
                time.sleep(1.0)

            return True

        except RuntimeError:
            raise
        except Exception as e:
            logger.error("执行查询导出流程失败: %s", e)
            import traceback
            traceback.print_exc()
            return False

    def export_report(self) -> None:
        pass
