from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Config:
    save_dir: Path = Path(r"D:\auto_reports")
    log_dir: Path = Path(__file__).resolve().parent / "logs"
    client_exe: Path = Path(r"D:\kaoqin\Main.exe")
    username: str = "制造工艺部"
    password: str = "123456"
    http_token: str = "18400021209"
    max_scan_retries: int = 2
    scan_retry_sleep_seconds: float = 0.5
