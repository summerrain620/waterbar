#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_refresh.py - 茶颜看板自动刷新脚本
自动查找最新xlsx -> 转换CSV -> 更新JSON -> 推送GitHub

使用方法:
  1. 自动模式: python auto_refresh.py
     (自动查找Downloads目录中最新的xlsx文件)
  2. 指定文件: python auto_refresh.py "C:/path/to/file.xlsx"
  3. 指定日期范围: python auto_refresh.py --all
     (处理Downloads目录中所有未处理的xlsx文件)
"""
import os, sys, csv, json, glob, subprocess, shutil, time
from datetime import datetime
from pathlib import Path

# ==================== 配置 ====================
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(REPO_DIR, "看板数据.json")
CONVERT_DIR = r"C:\Users\62398\.workbuddy\skills\chayan-xlsx-to-dashboard-csv\scripts"

# 直接导入 convert 模块（避免 subprocess 链条不稳定）
sys.path.insert(0, CONVERT_DIR)
import convert as _convert
XLSX_PATTERN = "门店销售报表_茶颜_各门店各商品销量数据明细_*.xlsx"
GITHUB_USERNAME = "summerrain620"
REPO_NAME = "waterbar"

# Read token from git-ignored file (do NOT hardcode tokens!)
def _load_token():
    token_path = os.path.join(REPO_DIR, ".github_token")
    if os.path.exists(token_path):
        with open(token_path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return None

GITHUB_TOKEN = _load_token()
if not GITHUB_TOKEN:
    raise RuntimeError("未找到 .github_token 文件，请确保该文件存在于仓库目录中")
GITHUB_PUSH_URL = f"https://{GITHUB_USERNAME}:{GITHUB_TOKEN}@github.com/{GITHUB_USERNAME}/{REPO_NAME}.git"
GITHUB_PAGES_URL = f"https://{GITHUB_USERNAME}.github.io/{REPO_NAME}/"
LOG_FILE = os.path.join(REPO_DIR, "auto_refresh.log")


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as lf:
            lf.write(datetime.now().strftime("%Y-%m-%d ") + line + "\n")
    except:
        pass


def find_latest_xlsx():
    """在Downloads目录找到最新的xlsx文件"""
    pattern = os.path.join(DOWNLOADS_DIR, XLSX_PATTERN)
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0]


def find_all_xlsx():
    """找到Downloads目录中所有xlsx文件，按时间排序"""
    pattern = os.path.join(DOWNLOADS_DIR, XLSX_PATTERN)
    files = glob.glob(pattern)
    if not files:
        return []
    files.sort(key=lambda f: os.path.getmtime(f))
    return files


def convert_xlsx_to_csv(xlsx_path, out_dir):
    """直接调用 convert 模块转换 xlsx → CSV（不走 subprocess，更稳定）"""
    # Determine output CSV path based on file naming convention
    fname = os.path.basename(xlsx_path)
    # Extract date from filename: ...YYYYMMDD_HH_MM_SS.xlsx → YYYYMMDD
    parts = fname.replace(".xlsx", "").split("_")
    date_str = None
    for p in parts:
        if len(p) == 8 and p.isdigit():
            date_str = p
            break
    if not date_str:
        log(f"无法从文件名提取日期: {fname}", "ERROR")
        return None

    csv_path = os.path.join(out_dir, f"茶颜日杯数_{date_str}.csv")
    try:
        store_count, _, _ = _convert.convert_one(xlsx_path, csv_path)
        log(f"  转换完成: {store_count} 条记录")
        return csv_path
    except Exception as e:
        log(f"转换失败: {e}", "ERROR")
        return None


def parse_csv_to_data(csv_path):
    """解析CSV，构建actualDaily和productDaily字典

    CSV格式: 日期,门店编码,杯数总计,冰茶销量,咖啡销量,罐子销量,冰茶试味,咖啡试味,罐子试味

    Returns: (date_str, actualDaily, productDaily)
        - date_str: 8位日期字符串，如 "20260706"
        - actualDaily: { "20260706": { "C02700058": 846, ... }, ... }
        - productDaily: { "20260706": { "C02700058": { icedTea, coffee, can, ... }, ... }, ... }
    """
    actualDaily = {}
    productDaily = {}
    date_str = None
    store_count = 0

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader, None)

        for row in reader:
            if len(row) < 3:
                continue

            date_raw = row[0].strip()
            store_code = row[1].strip()
            if not date_raw or not store_code:
                continue

            try:
                total = int(float(row[2])) if row[2] else 0
            except (ValueError, TypeError):
                total = 0

            # Normalize date: "2026-07-06" -> "20260706"
            date_norm = date_raw.replace("-", "")
            if len(date_norm) != 8:
                continue
            # Fix year: 2025 -> 2026 (dashboard convention)
            if date_norm.startswith("2025"):
                date_norm = "2026" + date_norm[4:]

            if not date_str:
                date_str = date_norm

            # Build actualDaily
            if date_norm not in actualDaily:
                actualDaily[date_norm] = {}
            actualDaily[date_norm][store_code] = total

            # Build productDaily (if 9-column format)
            if len(row) >= 9:
                try:
                    iced_tea = int(float(row[3])) if row[3] else 0
                    coffee = int(float(row[4])) if row[4] else 0
                    can = int(float(row[5])) if row[5] else 0
                    iced_tea_taste = int(float(row[6])) if row[6] else 0
                    coffee_taste = int(float(row[7])) if row[7] else 0
                    can_taste = int(float(row[8])) if row[8] else 0
                except (ValueError, TypeError):
                    iced_tea = coffee = can = 0
                    iced_tea_taste = coffee_taste = can_taste = 0

                if date_norm not in productDaily:
                    productDaily[date_norm] = {}
                productDaily[date_norm][store_code] = {
                    "icedTea": iced_tea,
                    "coffee": coffee,
                    "can": can,
                    "icedTeaTaste": iced_tea_taste,
                    "coffeeTaste": coffee_taste,
                    "canTaste": can_taste,
                }

            store_count += 1

    return date_str, actualDaily, productDaily, store_count


def update_json(date_str, actualDaily, productDaily):
    """更新看板数据.json，合并新数据"""
    with open(DATA_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Determine month key from date (20260706 -> "2026-07")
    month_key = f"{date_str[:4]}-{date_str[4:6]}"

    # Ensure months structure exists
    if "months" not in data:
        data["months"] = {}

    # If month doesn't exist, create from latest month's config
    if month_key not in data["months"]:
        existing_months = sorted(data["months"].keys())
        if existing_months:
            latest = existing_months[-1]
            new_month = dict(data["months"][latest])
            new_month["actualDaily"] = {}
            new_month["productDaily"] = {}
            data["months"][month_key] = new_month
            log(f"创建新月份: {month_key} (从 {latest} 复制配置)")
        else:
            log("JSON中没有任何月份数据！", "ERROR")
            return False

    month_data = data["months"][month_key]

    # Merge actualDaily (new data overwrites same date+store)
    if "actualDaily" not in month_data:
        month_data["actualDaily"] = {}
    merged_actual = 0
    for date, stores in actualDaily.items():
        if date not in month_data["actualDaily"]:
            month_data["actualDaily"][date] = {}
        month_data["actualDaily"][date].update(stores)
        merged_actual += len(stores)

    # Merge productDaily
    if "productDaily" not in month_data:
        month_data["productDaily"] = {}
    merged_prod = 0
    for date, stores in productDaily.items():
        if date not in month_data["productDaily"]:
            month_data["productDaily"][date] = {}
        month_data["productDaily"][date].update(stores)
        merged_prod += len(stores)

    # Update meta: 始终定位到最新月份，避免旧数据把页面拉回历史月份
    if "meta" not in data:
        data["meta"] = {}
    data["meta"]["months"] = sorted(data["months"].keys())
    data["meta"]["currentMonth"] = data["meta"]["months"][-1]  # 最新月份
    data["savedAt"] = datetime.now().isoformat()

    # Save JSON (minified for faster network transfer)
    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

    log(f"JSON更新完成: 月份={month_key}, 日期={date_str}")
    log(f"  杯数记录: {merged_actual} 条, 产品记录: {merged_prod} 条")
    return True


def _run_git(args, timeout=60):
    """可靠地运行 git 命令（适配 pythonw.exe 后台运行环境）

    问题背景：从 pythonw.exe（无控制台后台进程）调用 subprocess.run([git.exe, ...])
    会间歇性触发 [WinError 2] 系统找不到指定的文件（约 60% 概率）。
    根因是 Windows CreateProcess API 在无控制台上下文中的路径解析不稳定。

    解决方案：先尝试 list 模式（安全），失败后自动降级为 shell=True 模式（更可靠）。
    """
    GIT_EXE = r"C:\Users\62398\.workbuddy\vendor\PortableGit\mingw64\bin\git.exe"
    cmd = [GIT_EXE] + args

    # Attempt 1: list-based subprocess (preferred, no shell injection risk)
    try:
        return subprocess.run(cmd, cwd=REPO_DIR, capture_output=True,
                            text=True, encoding="utf-8", timeout=timeout)
    except FileNotFoundError:
        pass  # Fall through to shell=True fallback

    # Attempt 2: shell=True (more reliable from pythonw.exe background processes)
    try:
        shell_cmd = f'"{GIT_EXE}" ' + ' '.join(
            f'"{a}"' if (' ' in a or '&' in a) else a for a in args
        )
        return subprocess.run(shell_cmd, shell=True, cwd=REPO_DIR,
                            capture_output=True, text=True, encoding="utf-8", timeout=timeout)
    except Exception:
        raise


def git_push(commit_msg=None):
    """提交并推送到GitHub（带重试 + 自动解决冲突 + 兼容 pythonw 后台运行）"""
    MAX_RETRIES = 3
    RETRY_DELAY = 5  # seconds

    for attempt in range(MAX_RETRIES):
        try:
            # Step 0: 先拉取远程最新代码，避免推送冲突
            if attempt > 0:
                log(f"git 重试 (attempt {attempt+1}/{MAX_RETRIES})...")
            try:
                pull_result = _run_git(["pull", "--rebase", "origin", "main"])
                if pull_result.returncode != 0:
                    # Pull 失败不一定是致命的（可能是网络或已是最新）
                    stderr_tail = pull_result.stderr.strip().split('\n')[-1] if pull_result.stderr else ''
                    log(f"git pull 返回: {stderr_tail[:100]}", "WARN")
            except Exception as e:
                log(f"git pull 异常 (非致命): {e}", "WARN")

            # Step 1: Git add
            add_result = _run_git(["add", "看板数据.json"])
            if add_result.returncode != 0:
                log(f"git add 失败: {add_result.stderr}", "ERROR")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
                    continue
                return False

            # Step 2: Check if there are changes to commit
            diff_result = _run_git(["diff", "--cached", "--quiet"])
            if diff_result.returncode == 0:
                log("数据无变更，跳过推送")
                return True  # No changes is OK, not a failure

            # Step 3: Git commit
            if not commit_msg:
                commit_msg = f"数据更新: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            commit_result = _run_git(["commit", "-m", commit_msg])
            if commit_result.returncode != 0:
                log(f"git commit 失败: {commit_result.stderr}", "ERROR")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
                    continue
                return False
            log(f"已提交: {commit_msg}")

            # Step 4: Git push
            push_result = _run_git(["push", GITHUB_PUSH_URL, "main"])
            if push_result.returncode == 0:
                log("推送成功！GitHub Pages 将在1-2分钟内自动更新")
                return True

            # Push failed — check if it's a conflict we can resolve
            stderr = push_result.stderr or ''
            if "rejected" in stderr or "fetch first" in stderr:
                log("推送被拒(远程有新提交)，尝试 pull --rebase 后重推...")
                try:
                    _run_git(["pull", "--rebase", "origin", "main"])
                    push_result2 = _run_git(["push", GITHUB_PUSH_URL, "main"])
                    if push_result2.returncode == 0:
                        log("冲突解决成功，推送完成！")
                        return True
                except Exception as e:
                    log(f"冲突解决失败: {e}", "ERROR")

            log(f"推送失败: {stderr.strip()[-200:]}", "ERROR")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

        except FileNotFoundError as e:
            log(f"git.exe 不可用 (attempt {attempt+1}/{MAX_RETRIES}): {e}", "ERROR")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
        except Exception as e:
            log(f"git 异常 (attempt {attempt+1}/{MAX_RETRIES}): {e}", "ERROR")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

    log("推送失败：已达最大重试次数", "ERROR")
    return False


def process_one_xlsx(xlsx_path, do_push=True):
    """处理单个xlsx文件: 转换 -> 更新JSON -> 推送"""
    fname = os.path.basename(xlsx_path)
    log(f"处理: {fname}")

    # Step 1: Convert xlsx to CSV
    log("[1/3] 转换 xlsx -> CSV ...")
    csv_path = convert_xlsx_to_csv(xlsx_path, REPO_DIR)
    if not csv_path:
        return False
    log(f"  CSV: {os.path.basename(csv_path)}")

    # Step 2: Parse CSV and update JSON
    log("[2/3] 解析CSV，更新JSON ...")
    date_str, actualDaily, productDaily, store_count = parse_csv_to_data(csv_path)
    if not date_str:
        log("CSV解析失败！", "ERROR")
        return False
    log(f"  数据日期: {date_str}, 门店数: {store_count}")

    ok = update_json(date_str, actualDaily, productDaily)

    # Clean up CSV
    try:
        os.remove(csv_path)
    except:
        pass

    if not ok:
        return False

    # Step 3: Git push
    if do_push:
        log("[3/3] 推送到GitHub ...")
        git_push()
    else:
        log("[3/3] 跳过推送 (批量模式，最后统一推送)")

    return True


def main():
    print("=" * 50)
    print(f"  茶颜看板自动刷新  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    # Determine which xlsx files to process
    if len(sys.argv) > 1 and sys.argv[1] == "--all":
        # Process all xlsx files
        xlsx_files = find_all_xlsx()
        if not xlsx_files:
            log("Downloads目录中未找到xlsx文件", "ERROR")
            sys.exit(1)
        log(f"找到 {len(xlsx_files)} 个xlsx文件")
        success_count = 0
        for i, xlsx in enumerate(xlsx_files):
            log(f"--- [{i+1}/{len(xlsx_files)}] ---")
            if process_one_xlsx(xlsx, do_push=(i == len(xlsx_files) - 1)):
                success_count += 1
        log(f"完成: {success_count}/{len(xlsx_files)} 个文件处理成功")
    elif len(sys.argv) > 1 and sys.argv[1] != "--all":
        # Process specified file
        xlsx_path = sys.argv[1]
        if not os.path.exists(xlsx_path):
            log(f"文件不存在: {xlsx_path}", "ERROR")
            sys.exit(1)
        process_one_xlsx(xlsx_path, do_push=True)
    else:
        # Auto mode: find latest xlsx
        xlsx_path = find_latest_xlsx()
        if not xlsx_path:
            log("Downloads目录中未找到xlsx文件", "ERROR")
            log(f"  搜索目录: {DOWNLOADS_DIR}")
            log(f"  文件模式: {XLSX_PATTERN}")
            sys.exit(1)
        process_one_xlsx(xlsx_path, do_push=True)

    print()
    print(f"看板地址: {GITHUB_PAGES_URL}")
    print("=" * 50)


if __name__ == "__main__":
    main()
