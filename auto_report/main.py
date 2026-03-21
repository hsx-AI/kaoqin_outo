import argparse
import sys
import time
from pathlib import Path

from client_automation import ClientAutomation
from config import Config
from excel_handler import kill_all_excel_processes, save_all_open_workbooks
from utils import setup_logging


def _pre_cleanup(active_logger) -> None:
    """在自动化流程开始前清理残留的 Excel 进程，避免干扰 COM 检测"""
    killed = kill_all_excel_processes()
    if killed > 0:
        active_logger.info("预清理: 终止了 %d 个残留Excel进程", killed)
        time.sleep(2.0)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="auto_report")
    parser.add_argument("--client-exe", default=None)
    parser.add_argument("--no-launch-client", action="store_true")
    parser.add_argument("--save-excel", action="store_true")
    parser.add_argument("--save-dir", default=None)
    return parser.parse_args(argv)


def run_report_job(
    client_exe: str | None = None,
    save_dir: str | None = None,
    launch_client: bool = True,
    save_excel_when_no_launch: bool = False,
    logger=None,
    global_timeout_seconds: float = 180.0,
) -> tuple[int, list[Path]]:
    config = Config()
    active_logger = setup_logging(config.log_dir) if logger is None else logger
    automation: ClientAutomation | None = None

    effective_client_exe = str(config.client_exe) if client_exe is None else client_exe
    effective_save_dir = config.save_dir if save_dir is None else type(config.save_dir)(save_dir)
    effective_save_dir.mkdir(parents=True, exist_ok=True)

    job_deadline = time.time() + global_timeout_seconds

    if launch_client:
        _pre_cleanup(active_logger)

        try:
            automation = ClientAutomation(executable_path=effective_client_exe)
            hwnd = automation.launch()
            if hwnd:
                active_logger.info("客户端已启动，窗口句柄: %s", hwnd)

                if time.time() > job_deadline:
                    active_logger.error("全局超时: 登录前")
                    return 7, []

                if automation.login(username=config.username, password=config.password):
                    active_logger.info("登录操作执行完成")

                    if time.time() > job_deadline:
                        active_logger.error("全局超时: 查询前")
                        return 7, []

                    active_logger.info("正在查找并点击[原始数据查询]...")
                    if automation.query_original_data():
                        active_logger.info("成功点击[原始数据查询]")

                        if time.time() > job_deadline:
                            active_logger.error("全局超时: 导出前")
                            return 7, []

                        active_logger.info("开始执行查询和导出流程...")
                        remaining = max(30.0, job_deadline - time.time())
                        try:
                            if automation.perform_query_and_export(timeout_seconds=remaining):
                                active_logger.info("数据查询与导出流程完成，开始保存Excel报表...")
                                results = save_all_open_workbooks(
                                    save_dir=effective_save_dir,
                                    max_scan_retries=config.max_scan_retries,
                                    scan_retry_sleep_seconds=config.scan_retry_sleep_seconds,
                                    logger=active_logger,
                                )

                                if not results:
                                    active_logger.error("未找到或保存任何Excel工作簿")
                                    return 2, []
                                else:
                                    active_logger.info("Excel报表保存成功")
                                    return 0, results
                            else:
                                active_logger.error("数据查询与导出流程失败")
                                return 6, []
                        except RuntimeError as e:
                            active_logger.error("自动化流程异常中断: %s", e)
                            return 8, []
                    else:
                        active_logger.error("未找到[原始数据查询]按钮或点击失败")
                        return 5, []
                else:
                    active_logger.error("未找到登录窗口或登录操作失败")
                    return 4, []
            else:
                active_logger.warning("客户端已启动，但未检测到可见窗口")
        except Exception as e:
            active_logger.error("启动客户端失败: %s", e)
            return 3, []
        finally:
            if automation is not None:
                if automation.close_process_by_name("XDKQ_HEC"):
                    active_logger.info("已关闭进程: XDKQ_HEC")
                else:
                    active_logger.warning("未关闭进程或进程不存在: XDKQ_HEC")
                if automation.close_image_process("EXCEL"):
                    active_logger.info("已关闭进程: EXCEL")
                else:
                    active_logger.warning("未关闭进程或进程不存在: EXCEL")

    if not save_excel_when_no_launch:
        return 0, []

    only_save_results = save_all_open_workbooks(
        save_dir=effective_save_dir,
        max_scan_retries=config.max_scan_retries,
        scan_retry_sleep_seconds=config.scan_retry_sleep_seconds,
        logger=active_logger,
    )

    if not only_save_results:
        return 2, []
    return 0, only_save_results


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    config = Config()
    logger = setup_logging(config.log_dir)
    code, _results = run_report_job(
        client_exe=args.client_exe,
        save_dir=args.save_dir,
        launch_client=not args.no_launch_client,
        save_excel_when_no_launch=args.save_excel,
        logger=logger,
    )
    return code


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
