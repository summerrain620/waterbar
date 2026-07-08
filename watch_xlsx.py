#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
watch_xlsx.py - 实时监控下载文件夹，xlsx 文件一到立即自动刷新

启动方式：
  python watch_xlsx.py         前台运行（终端可见）
  pythonw watch_xlsx.py        后台静默运行（无窗口）

开机自启：
  已配置计划任务 "ChayanDashboardWatch"（用户登录时启动）
"""

import os, sys, glob, time, json, subprocess, hashlib
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ==================== 配置 ====================
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(REPO_DIR, "watch_xlsx.log")
STATE_FILE = os.path.join(REPO_DIR, ".watch_state.json")
XLSX_PATTERN = "门店销售报表_茶颜_各门店各商品销量数据明细_*.xlsx"
DEBOUNCE_SECONDS = 10  # 文件稳定后等待秒数再处理
AUTO_REFRESH_SCRIPT = os.path.join(REPO_DIR, "auto_refresh.py")
PYTHON_EXE = r"C:\Users\62398\.workbuddy\binaries\python\envs\default\Scripts\python.exe"


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as lf:
            lf.write(line + "\n")
    except:
        pass


def file_hash(path):
    """快速计算文件哈希，用于判断文件是否真的变了"""
    try:
        stat = os.stat(path)
        return f"{stat.st_size}_{stat.st_mtime}"
    except:
        return None


def load_state():
    """加载上次处理记录"""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False)


def run_refresh(xlsx_path):
    """调用 auto_refresh.py 处理指定 xlsx"""
    log(f"开始处理: {os.path.basename(xlsx_path)}")
    cmd = [PYTHON_EXE, AUTO_REFRESH_SCRIPT, xlsx_path]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.returncode == 0:
        log(f"✅ 处理成功")
        # 打印关键输出
        for line in result.stdout.split("\n"):
            if any(kw in line for kw in ["处理:", "数据日期", "门店数", "JSON更新", "推送", "成功", "失败"]):
                log(f"  {line.strip()}")
    else:
        log(f"❌ 处理失败 (exit={result.returncode})", "ERROR")
        if result.stderr:
            for line in result.stderr.strip().split("\n")[-5:]:
                log(f"  {line}", "ERROR")
    return result.returncode == 0


class XlsxHandler(FileSystemEventHandler):
    def __init__(self):
        self.pending = {}  # {path: timer}

    def _should_process(self, path):
        fname = os.path.basename(path)
        # 只处理 xlsx
        if not fname.endswith(".xlsx"):
            return False
        # 只处理茶颜门店销售报表
        if not fname.startswith("门店销售报表_茶颜"):
            return False
        # 排除临时文件
        if fname.startswith("~$"):
            return False
        return True

    def _check_and_trigger(self, path):
        """防抖检查，文件稳定后才触发"""
        if not self._should_process(path):
            return

        fh = file_hash(path)
        if not fh:
            return

        state = load_state()
        last = state.get("last_processed", {})
        if last.get("hash") == fh and last.get("path") == path:
            log(f"跳过重复: {os.path.basename(path)} (文件未变化)")
            return

        # 确保文件不再被写入（再等一次确认）
        time.sleep(DEBOUNCE_SECONDS)
        fh2 = file_hash(path)
        if fh2 != fh:
            log(f"文件还在变化中，跳过: {os.path.basename(path)}")
            return

        # 处理
        log(f"检测到文件就绪: {os.path.basename(path)}")
        if run_refresh(path):
            state["last_processed"] = {"path": path, "hash": fh2, "time": datetime.now().isoformat()}
            save_state(state)

    def on_created(self, event):
        if not event.is_directory:
            path = event.src_path
            log(f"新文件创建: {os.path.basename(path)}")
            # 延迟检查，等文件写完整
            from threading import Timer
            Timer(DEBOUNCE_SECONDS, self._check_and_trigger, args=[path]).start()

    def on_modified(self, event):
        if not event.is_directory:
            path = event.src_path
            self._check_and_trigger(path)

    def on_moved(self, event):
        if not event.is_directory:
            dest = event.dest_path
            log(f"文件移动: {os.path.basename(dest)}")
            self._check_and_trigger(dest)


def main():
    log("=" * 50)
    log("  茶颜看板 XLSX 实时监控 启动")
    log(f"  监控目录: {DOWNLOADS_DIR}")
    log(f"  文件模式: {XLSX_PATTERN}")
    log("=" * 50)

    # 启动时检查是否已有未处理的新文件
    pattern = os.path.join(DOWNLOADS_DIR, XLSX_PATTERN)
    existing = glob.glob(pattern)
    if existing:
        existing.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        latest = existing[0]
        state = load_state()
        fh = file_hash(latest)
        last = state.get("last_processed", {})
        if last.get("hash") != fh:
            mtime = datetime.fromtimestamp(os.path.getmtime(latest))
            age_mins = (datetime.now() - mtime).total_seconds() / 60
            if age_mins > 2:  # 启动时只处理已经存在超过 2 分钟的文件
                log(f"启动时发现未处理文件: {os.path.basename(latest)} ({age_mins:.0f}分钟前)")
                handler = XlsxHandler()
                handler._check_and_trigger(latest)

    # 启动文件监控
    observer = Observer()
    handler = XlsxHandler()
    observer.schedule(handler, DOWNLOADS_DIR, recursive=False)
    observer.start()
    log("监控已启动，等待 xlsx 文件...")

    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        log("收到停止信号，关闭监控")
        observer.stop()
    observer.join()
    log("监控已停止")


if __name__ == "__main__":
    main()
