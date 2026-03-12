from dataclasses import dataclass
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

@dataclass
class ClientAutomation:
    executable_path: str | None = None
    pid: int | None = None
    hwnd: int | None = None
    process: subprocess.Popen | None = None

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
                    capture_output=True,
                    text=True,
                    check=False,
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
                    capture_output=True,
                    text=True,
                    check=False,
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
                    capture_output=True,
                    text=True,
                    check=False,
                )
                if result.returncode == 0:
                    return True
            except Exception:
                continue
        return False

    def _set_clipboard_text(self, text: str) -> None:
        """Helper to set text to clipboard for unicode support"""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except Exception:
            # Try to close if something went wrong
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            raise

    def login(self, username: str, password: str, timeout_seconds: float = 10.0) -> bool:
        """
        查找标题包含“登录”的窗口，并尝试输入用户名密码
        注意：这里假设登录窗口属于当前启动的进程
        """
        if not self.pid:
            raise RuntimeError("Process not launched")

        # 1. 寻找登录窗口
        login_hwnd: int | None = None
        deadline = time.time() + float(timeout_seconds)

        def on_window(hwnd: int, _lparam: int) -> bool:
            nonlocal login_hwnd
            try:
                # 检查进程ID
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if int(window_pid) != int(self.pid):
                    return True
                
                # 检查可见性
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                # 检查标题
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
            # 如果没找到特定标题的窗口，尝试使用主窗口（可能就是主窗口登录）
            if self.hwnd and win32gui.IsWindowVisible(self.hwnd):
                login_hwnd = self.hwnd
            else:
                return False

        # 增加等待时间，确保页面加载完成
        time.sleep(5.0)

        # 2. 激活窗口并输入
        try:
            # 尝试将窗口前置
            if win32gui.IsIconic(login_hwnd):
                win32gui.ShowWindow(login_hwnd, 9)  # SW_RESTORE
            win32gui.SetForegroundWindow(login_hwnd)
            time.sleep(0.5)

            shell = win32com.client.Dispatch("WScript.Shell")
            
            # 输入用户名
            # 先清除可能存在的默认值 (Ctrl+A -> Del)
            shell.SendKeys("^a")
            time.sleep(0.1)
            shell.SendKeys("{DELETE}")
            time.sleep(0.1)
            
            # 使用剪贴板输入中文用户名
            self._set_clipboard_text(username)
            time.sleep(0.2)
            shell.SendKeys("^v")
            time.sleep(0.2)
            
            # 切换到密码框
            shell.SendKeys("{TAB}")
            time.sleep(0.2)
            
            # 输入密码
            shell.SendKeys(password)
            time.sleep(0.2)
            
            # 提交
            shell.SendKeys("{ENTER}")
            time.sleep(1.0)
            
            return True
        except Exception:
            return False

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

    def query_original_data(self, timeout_seconds: float = 10.0) -> bool:
        """
        查找并点击“原始数据查询”按钮
        由于是.NET WinForms程序，使用uia backend
        """
        if not self.pid:
            raise RuntimeError("Process not launched")

        # 定义尝试点击的逻辑
        def try_click_with_backend(backend: str) -> bool:
            try:
                # WinForms程序通常需要uia backend
                app = Application(backend=backend).connect(process=self.pid)
                
                # 获取主窗口
                # 根据dump结果，窗口类名是WindowsForms10.Window.8.app...
                # 我们可以尝试匹配标题
                windows = app.windows(title_re=".*先达考勤管理系统.*")
                
                for window in windows:
                    # 将窗口前置
                    try:
                        window.set_focus()
                    except Exception:
                        pass

                    # 策略1：查找ToolStrip1中的按钮
                    # 结构显示有 'ToolStrip1'
                    try:
                        # 在UIA模式下，ToolStrip通常包含Button子元素
                        # 查找所有名为"原始数据查询"的元素
                        btns = window.descendants(title="原始数据查询", control_type="Button")
                        for btn in btns:
                            try:
                                btn.click_input()
                                return True
                            except Exception:
                                pass
                        
                        # 如果上面的没找到，尝试模糊匹配
                        btns = window.descendants(title_re=".*原始数据查询.*", control_type="Button")
                        for btn in btns:
                            try:
                                btn.click_input()
                                return True
                            except Exception:
                                pass
                    except Exception:
                        pass

                    # 策略2：通过菜单点击 "数据查询" -> "原始数据查询"
                    # 结构显示有 'MenuStrip1'
                    try:
                        # 查找菜单项
                        menu_item = window.descendants(title="数据查询", control_type="MenuItem")
                        if menu_item:
                            menu_item[0].click_input()
                            time.sleep(0.5)
                            # 查找子菜单
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
            # 根据Debug结果，这是WinForms程序，强烈建议使用uia
            if try_click_with_backend("uia"):
                return True
            
            time.sleep(1.0)
            
        # 如果超时仍未找到，导出UI结构用于调试
        print("未找到按钮，正在导出UI结构到 ui_dump_uia.txt ...")
        self._dump_debug_info("uia", "ui_dump_uia.txt")
        
        return False

    def perform_query_and_export(self) -> bool:
        """
        执行查询并导出数据流程：
        1. 点击“查询”按钮
        2. 在弹出的查询窗口中点击“确定”
        3. 等待10秒加载
        4. 点击“结果导出”
        5. 在确认弹窗中点击“是”
        6. 等待20秒导出
        """
        try:
            app = Application(backend="uia").connect(process=self.pid)
            
            # 1. 点击“查询”按钮 (在主界面)
            print("正在查找主界面“查询”按钮...")
            main_window = app.window(title_re=".*先达考勤管理系统.*")
            main_window.set_focus()
            
            # 查找主界面中的“查询”按钮 (注意：不是顶部工具栏的，是内容区域的)
            query_btn = main_window.descendants(title="查询", control_type="Button")
            if not query_btn:
                print("未找到主界面“查询”按钮")
                return False
            
            query_btn[0].click_input()
            print("点击主界面“查询”按钮")
            time.sleep(5)

            # 2. 在弹出的“查询”窗口中点击“确定”
            print("正在等待“查询”窗口...")
            # 使用 Desktop 查找该进程的所有窗口，这比 app.windows() 更可靠
            query_dialog = None
            for i in range(10):  # 尝试10次，每次1.0秒
                try:
                    windows = Desktop(backend="uia").windows(process=self.pid, visible_only=True)
                    if i == 0:
                        print(f"当前可见窗口: {[w.window_text() for w in windows]}")
                        
                    for w in windows:
                        if "查询" in w.window_text():
                            query_dialog = w
                            break
                    if query_dialog:
                        break
                except Exception as e:
                    print(f"查找窗口出错: {e}")
                time.sleep(1.0)

            if query_dialog:
                try:
                    query_dialog.set_focus()
                    confirm_btn = query_dialog.descendants(title="确定", control_type="Button")
                    if confirm_btn:
                        confirm_btn[0].click_input()
                        print("点击查询窗口“确定”按钮")
                    else:
                        print("未找到查询窗口“确定”按钮，尝试按回车键...")
                        query_dialog.type_keys("{ENTER}")
                except Exception as e:
                    print(f"操作查询窗口失败: {e}，尝试按回车键...")
                    try:
                        query_dialog.type_keys("{ENTER}")
                    except:
                        pass
            else:
                print("未找到查询窗口(超时)，尝试在主窗口按回车键...")
                try:
                    main_window.type_keys("{ENTER}")
                except:
                    pass

            # 3. 等待5S等他加载完
            print("等待5秒数据加载...")
            time.sleep(5 )

            # 4. 点击“结果导出”
            print("正在查找“结果导出”按钮...")

            main_wrapper = None
            try:
                main_window = app.window(title_re=".*先达考勤.*")
                if not main_window.exists(timeout=5):
                    print("未找到主窗口，尝试重新连接应用...")
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
                    print(f"已定位主窗口: {main_wrapper.window_text()}")
                except Exception as e:
                    print(f"主窗口set_focus失败(忽略): {e}")
            except Exception as e:
                print(f"重新查找主窗口失败: {e}")
                return False

            parent_hwnd = int(self.hwnd) if self.hwnd else 0
            if parent_hwnd <= 0 and self.pid:
                temp_hwnd = self._find_visible_window_for_pid(int(self.pid))
                if temp_hwnd:
                    parent_hwnd = int(temp_hwnd)
                    self.hwnd = parent_hwnd

            if parent_hwnd <= 0:
                print("未找到主窗口句柄，无法执行Win32相邻按钮点击")
                return False

            if not self._click_export_by_query_neighbor(parent_hwnd):
                print("Win32相邻按钮点击未命中")
                return False

            print("已执行Win32相邻按钮点击")
            
            # 5. 会有弹窗是否导出 点击是 (或回车)
            print("等待3秒确认弹窗出现...")
            time.sleep(3)
            
            print("正在确认导出(发送回车)...")
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys("{ENTER}")
            except Exception:
                try:
                    main_window.type_keys("{ENTER}")
                except Exception:
                    pass
            
            # 6. 等待导出完成
            print("等待15秒导出完成...")
            time.sleep(15)
            
            return True

        except Exception as e:
            print(f"执行查询导出流程失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def export_report(self) -> None:
        pass
