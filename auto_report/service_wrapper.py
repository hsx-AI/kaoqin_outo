"""
考勤报表后台服务包装器

功能：
  - 无窗口后台运行（配合 pythonw.exe）
  - Windows 命名互斥量保证单实例
  - HTTP 服务崩溃后自动重启
  - PID 文件方便外部管理
  - 支持 --stop / --status 命令

使用方式：
  pythonw.exe service_wrapper.py                # 无窗口后台运行
  python      service_wrapper.py                # 有控制台运行（调试用）
  python      service_wrapper.py --stop         # 停止服务
  python      service_wrapper.py --status       # 查看状态
"""

import ctypes
import logging
import os
import signal
import subprocess
import sys
import threading
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

MUTEX_NAME = "Global\\KaoqinAutoReportService"
PID_FILE = SCRIPT_DIR / "service.pid"
DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 6648
RESTART_DELAY_SECONDS = 5

logger = logging.getLogger("auto_report")


# ── 单实例互斥量 ──────────────────────────────────────────────────────────

def _acquire_mutex():
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        kernel32.CloseHandle(mutex)
        return None
    return mutex


def _release_mutex(mutex):
    if mutex:
        try:
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)
        except Exception:
            pass


# ── PID 文件 ──────────────────────────────────────────────────────────────

def _write_pid():
    try:
        PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        pass


def _remove_pid():
    try:
        PID_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def _read_pid() -> int | None:
    try:
        return int(PID_FILE.read_text(encoding="utf-8").strip())
    except Exception:
        return None


def _is_process_alive(pid: int) -> bool:
    try:
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
    except Exception:
        pass
    return False


# ── 命令：停止 / 状态 ────────────────────────────────────────────────────

def cmd_stop():
    pid = _read_pid()
    if pid and _is_process_alive(pid):
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                capture_output=True, text=True, check=False,
            )
            print(f"已停止服务 (PID: {pid})")
        except Exception as e:
            print(f"停止失败: {e}")
    else:
        print("服务未在运行")
    _remove_pid()


def cmd_status():
    pid = _read_pid()
    if pid and _is_process_alive(pid):
        print(f"服务正在运行 (PID: {pid})")
    else:
        print("服务未在运行")
        if pid:
            _remove_pid()


# ── 主服务循环 ────────────────────────────────────────────────────────────

def cmd_run():
    mutex = _acquire_mutex()
    if mutex is None:
        pid = _read_pid()
        msg = f"服务已在运行 (PID: {pid})" if pid else "另一个实例已在运行"
        print(msg)
        return 1

    _write_pid()

    from config import Config
    from utils import setup_logging

    config = Config()
    svc_logger = setup_logging(config.log_dir)
    svc_logger.info("后台服务启动 (PID: %d)", os.getpid())

    running = True
    server_ref = [None]

    def on_signal(signum, _frame):
        nonlocal running
        running = False
        svc_logger.info("收到信号 %s, 正在停止...", signum)
        srv = server_ref[0]
        if srv is not None:
            threading.Thread(target=srv.shutdown, daemon=True).start()

    signal.signal(signal.SIGINT, on_signal)
    signal.signal(signal.SIGTERM, on_signal)

    try:
        _register_console_ctrl_handler(on_signal, svc_logger)
    except Exception:
        pass

    from http_service import ReportHTTPServer, ReportHandler

    token = os.getenv("AUTO_REPORT_TOKEN") or config.http_token
    if not token:
        svc_logger.error("HTTP token 未配置")
        _remove_pid()
        _release_mutex(mutex)
        return 1

    try:
        while running:
            try:
                server = ReportHTTPServer(
                    (DEFAULT_HOST, DEFAULT_PORT),
                    ReportHandler,
                    logger=svc_logger,
                    client_exe=None,
                    save_dir=None,
                    access_token=token,
                )
                server_ref[0] = server
                svc_logger.info(
                    "HTTP服务已启动: http://%s:%s", DEFAULT_HOST, DEFAULT_PORT
                )

                server_thread = threading.Thread(
                    target=server.serve_forever, daemon=True
                )
                server_thread.start()

                while running and server_thread.is_alive():
                    time.sleep(1.0)

                server.shutdown()
                server_ref[0] = None

                if not running:
                    break

                svc_logger.warning(
                    "HTTP服务意外退出, %d秒后重启...", RESTART_DELAY_SECONDS
                )
                time.sleep(RESTART_DELAY_SECONDS)

            except OSError as e:
                svc_logger.error("端口绑定失败: %s, %d秒后重试...", e, RESTART_DELAY_SECONDS)
                time.sleep(RESTART_DELAY_SECONDS)
            except Exception as e:
                svc_logger.error(
                    "服务异常: %s, %d秒后重启...", e, RESTART_DELAY_SECONDS
                )
                time.sleep(RESTART_DELAY_SECONDS)
    finally:
        _remove_pid()
        _release_mutex(mutex)
        svc_logger.info("后台服务已停止")

    return 0


def _register_console_ctrl_handler(on_signal_fn, svc_logger):
    """注册 Windows 控制台事件处理（关机/注销时优雅退出）"""
    import ctypes
    from ctypes import wintypes

    CTRL_C_EVENT = 0
    CTRL_BREAK_EVENT = 1
    CTRL_CLOSE_EVENT = 2
    CTRL_LOGOFF_EVENT = 5
    CTRL_SHUTDOWN_EVENT = 6

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)
    def handler(event):
        if event in (CTRL_CLOSE_EVENT, CTRL_LOGOFF_EVENT, CTRL_SHUTDOWN_EVENT,
                     CTRL_C_EVENT, CTRL_BREAK_EVENT):
            svc_logger.info("收到Windows控制事件 %d, 正在停止...", event)
            on_signal_fn(event, None)
            time.sleep(3)
            return True
        return False

    ctypes.windll.kernel32.SetConsoleCtrlHandler(handler, True)


# ── 入口 ──────────────────────────────────────────────────────────────────

def main():
    os.chdir(str(SCRIPT_DIR))

    if len(sys.argv) > 1:
        arg = sys.argv[1].lower().lstrip("-")
        if arg == "stop":
            cmd_stop()
            return 0
        if arg == "status":
            cmd_status()
            return 0
        if arg not in ("run", "start"):
            print(f"未知命令: {sys.argv[1]}")
            print("用法: service_wrapper.py [--stop | --status]")
            return 1

    return cmd_run()


if __name__ == "__main__":
    raise SystemExit(main())
