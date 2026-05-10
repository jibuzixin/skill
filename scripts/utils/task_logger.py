import subprocess
import atexit
import os
import sys
import logging
import socket
import time
import shutil
from pathlib import Path

THIS_FILE = Path(__file__).absolute()
THIS_DIR = THIS_FILE.parent
SCRIPT_PATH = THIS_DIR / "single_task_recorder.py"

logger = logging.getLogger(__name__)

# 全局单例：只启动 1 次 mitm
mitm_process = None
mitm_port = None

# 自动获取空闲端口
def find_free_port(start=10000, end=50000):
    for port in range(start, end):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
            return port
        except Exception:
            continue
    return 8083

def start_mitm_once() -> bool:
    global mitm_process, mitm_port
    if mitm_process is not None:
        return True

    port = find_free_port()

    # 直接用当前环境的 mitmdump 可执行文件
    mitmdump_path = shutil.which("mitmdump")
    if mitmdump_path is None:
        logger.warning(f"⚠️ 未能找到可用的 mitmdump 程序，取消 --mitm 参数使用")
        return False
    mitmdump_path = str(mitmdump_path)

    cmd = [
        mitmdump_path,
        "-s", str(SCRIPT_PATH),
        "-p", str(port),
        "-q"
    ]

    mitm_process = subprocess.Popen(
        cmd,
        cwd=str(THIS_DIR)
    )
    mitm_port = port
    time.sleep(1.0)
    logger.info(f"✅ mitm 正向代理启动成功: 127.0.0.1:{port}")
    return True

# ------------------------------
# 设置日志路径：写入端口绑定的临时文件（替代共享内存）
# ------------------------------
def set_current_log_path(log_path):
    if not mitm_port:
        return
    sync_file = THIS_DIR / f"mitm_log_path_{mitm_port}.tmp"
    try:
        sync_file.write_text(str(log_path), encoding="utf-8")
        logger.info(f"📝 设置大模型请求日志保存路径为：{log_path}")
    except Exception as e:
        logger.error(f"📝 设置大模型请求日志路径错误：{str(e)}")
        pass

# 获取当前 A 程序的代理环境
def get_proxy_env() -> dict | None:
    if mitm_port is None:
        return None
    env = os.environ.copy()
    proxy = f"http://127.0.0.1:{mitm_port}"
    env["HTTP_PROXY"] = proxy
    env["HTTPS_PROXY"] = proxy
    env["NO_PROXY"] = ""
    return env

# 退出时清理
def stop_mitm():
    global mitm_process, mitm_port

    if mitm_process:
        try:
            import os
            import signal
            os.kill(mitm_process.pid, signal.SIGKILL)
        except:
            try:
                os.kill(mitm_process.pid, signal.SIGKILL)  # 强杀
            except:
                pass
        mitm_process = None

    # 清理端口对应的同步文件
    if mitm_port:
        sync_file = THIS_DIR / f"mitm_log_path_{mitm_port}.tmp"
        try:
            if sync_file.exists():
                sync_file.unlink()
        except Exception:
            pass

    mitm_port = None
    logger.info(f"🛑 已停止 mitm 并清理同步文件")

atexit.register(stop_mitm)