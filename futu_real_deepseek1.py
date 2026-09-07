#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FutuOpenD A股/港股/美股自动交易 (市价单版)
版本: v2.3
修复: OpenSecTradeContext 不支持 acc_id 和 is_simulate，改为方法参数传递
功能：
- 模拟盘真实提交模拟订单
- A股仅允许模拟盘，禁止实盘
- 自动识别市场时区，非交易时段休眠，开盘前10分钟唤醒
- 滚动日志（20MB/个，保留5个）
- 钉钉告警、崩溃自动重启
- 多层次价格网格买入，动态止盈
"""
import base64
import csv
import hmac
import hashlib
import json
import os
import sys
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Tuple

import futu as ft
from dotenv import load_dotenv
import openpyxl
import logging
import logging.handlers

load_dotenv()

# ==================== 环境变量容错解析 ====================
def _env_int(key: str, default: int) -> int:
    raw = os.getenv(key)
    if raw is None or raw.strip() == "":
        return default
    cleaned = raw.strip().strip("|").strip().strip('"\'')
    try:
        return int(cleaned)
    except ValueError:
        print(f"[配置] 环境变量 {key} 值无效: '{raw}'，已使用默认值 {default}")
        return default

def _env_float(key: str, default: float) -> float:
    raw = os.getenv(key)
    if raw is None or raw.strip() == "":
        return default
    cleaned = raw.strip().strip("|").strip().strip('"\'')
    try:
        return float(cleaned)
    except ValueError:
        print(f"[配置] 环境变量 {key} 值无效: '{raw}'，已使用默认值 {default}")
        return default

def _env_bool(key: str, default: bool) -> bool:
    raw = os.getenv(key)
    if raw is None:
        return default
    return raw.strip().lower() in ("true", "1", "yes", "是")

# ==================== 日志系统 ====================
def setup_logging():
    log_level = os.getenv("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, log_level, logging.INFO)
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "futu_autoTrade_data")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "futu_autoTrade_run.log")

    root = logging.getLogger()
    root.setLevel(level)
    for h in root.handlers[:]:
        root.removeHandler(h)

    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    max_bytes = _env_int("LOG_MAX_BYTES", 20 * 1024 * 1024)
    backup_count = _env_int("LOG_BACKUP_COUNT", 5)
    file_handler = logging.handlers.RotatingFileHandler(
        log_file, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8"
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    root.addHandler(console_handler)

    # 钉钉告警
    ding_webhook = os.getenv("DINGTALK_WEBHOOK_futu")
    if ding_webhook:
        class DingTalkHandler(logging.Handler):
            def __init__(self, webhook, secret=None):
                super().__init__()
                self.webhook = webhook
                self.secret = secret
                self.setLevel(logging.ERROR)
            def emit(self, record):
                try:
                    msg = self.format(record)
                    push_to_dingtalk(self.webhook, self.secret, f"🚨 系统异常\n\n{msg}")
                except Exception:
                    pass
        ding_handler = DingTalkHandler(ding_webhook, os.getenv("DINGTALK_SECRET_futu"))
        ding_handler.setFormatter(formatter)
        root.addHandler(ding_handler)

    logging.info("=" * 60)
    logging.info(f"日志系统初始化完成，级别: {log_level}，文件: {log_file}")

def push_to_dingtalk(webhook: str, secret: str = None, text: str = None) -> bool:
    if not webhook:
        return False
    try:
        timestamp = str(int(time.time() * 1000))
        if secret:
            sign_string = f"{timestamp}\n{secret}".encode("utf-8")
            sign = base64.b64encode(hmac.new(secret.encode("utf-8"), sign_string, hashlib.sha256).digest()).decode("utf-8")
            url = webhook + f"&timestamp={timestamp}&sign={sign}"
        else:
            url = webhook
        payload = {"msgtype": "text", "text": {"content": text}}
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                     headers={"Content-Type": "application/json", "User-Agent": "Python"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8")).get("errcode") in (0, "0", None)
    except Exception as e:
        logging.error(f"钉钉推送失败: {e}")
        return False

# ==================== 交易时间管理 ====================
def get_market_timezone_offset(market, dt=None):
    if dt is None:
        dt = datetime.utcnow()
    if market in ('CN', 'HK'):
        return 8
    elif market == 'US':
        year = dt.year
        def second_sunday(month):
            first = datetime(year, month, 1)
            days_to_sunday = (6 - first.weekday()) % 7
            return first + timedelta(days=days_to_sunday + 7)
        dst_start = second_sunday(3)
        dst_end = second_sunday(11)
        if dst_start <= dt.replace(tzinfo=None) < dst_end:
            return -4
        else:
            return -5
    return 0

def is_weekday(dt):
    return dt.weekday() < 5

def is_market_open(market, dt_utc):
    offset = get_market_timezone_offset(market, dt_utc)
    local_dt = dt_utc + timedelta(hours=offset)
    if not is_weekday(local_dt):
        return False
    hour = local_dt.hour
    minute = local_dt.minute
    if market == 'CN':
        close_hour = 15
    elif market == 'HK':
        close_hour = 16
    elif market == 'US':
        close_hour = 16
    else:
        return False
    if hour > 9 or (hour == 9 and minute >= 30):
        if hour < close_hour or (hour == close_hour and minute == 0):
            return True
    return False

def get_next_market_open(markets, from_dt_utc=None):
    if from_dt_utc is None:
        from_dt_utc = datetime.utcnow()
    for m in markets:
        if is_market_open(m, from_dt_utc):
            return None
    for days in range(7):
        for hour in range(24):
            for minute in (0, 30):
                dt_candidate = from_dt_utc + timedelta(days=days, hours=hour, minutes=minute)
                if dt_candidate <= from_dt_utc:
                    continue
                for m in markets:
                    if is_market_open(m, dt_candidate):
                        return dt_candidate
    return None

def wait_until_market_open(markets):
    while True:
        now = datetime.utcnow()
        next_open = get_next_market_open(markets, now)
        if next_open is None:
            return
        wake_time = next_open - timedelta(minutes=10)
        if wake_time <= now:
            return
        sleep_sec = (wake_time - now).total_seconds()
        if sleep_sec > 0:
            logging.info(f"非交易时段，休眠 {sleep_sec/60:.1f} 分钟，将于 {wake_time.strftime('%Y-%m-%d %H:%M:%S')} UTC 唤醒")
            time.sleep(sleep_sec)

# ==================== 配置加载 ====================
def load_config() -> Tuple[Dict, Dict]:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    files = {
        "HK": os.path.join(base_dir, "hk.xlsx"),
        "US": os.path.join(base_dir, "us.xlsx"),
        "CN": os.path.join(base_dir, "a.xlsx"),
    }
    config = {}
    for market, filepath in files.items():
        if not os.path.exists(filepath):
            logging.warning(f"股票池文件不存在: {filepath}，跳过")
            continue
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            for row in range(2, ws.max_row + 1):
                code = ws.cell(row=row, column=1).value
                ref_high = ws.cell(row=row, column=2).value
                ref_low = ws.cell(row=row, column=3).value
                enabled = ws.cell(row=row, column=4).value
                if code and enabled and str(enabled).upper() in ("YES", "是", "1"):
                    try:
                        ref_high = float(ref_high)
                        ref_low = float(ref_low)
                    except (TypeError, ValueError):
                        continue
                    config[code] = {"ref_high": ref_high, "ref_low": ref_low}
        except Exception as e:
            logging.error(f"读取 {filepath} 失败: {e}")

    system_params = {
        "buy_ratio": _env_float("BUY_RATIO", 0.01),
        "cooldown_days": _env_int("COOLDOWN_DAYS", 7),
        "profit_trigger": _env_float("PROFIT_TRIGGER", 0.05),
        "stop_loss": _env_float("STOP_LOSS", 0.015),
        "profit_step": 0.01,
        "break_alert_max": 4,
        "break_alert_interval": 1020,
        "push_interval": 3.0,
        "data_source": "FUTU",
    }
    return config, system_params

# ==================== CSV辅助函数 ====================
PROGRAM_NAME = "futu_autoTrade"
VERSION = "v2.3"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_data")
os.makedirs(DATA_DIR, exist_ok=True)

TRADE_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_tradeRecord.csv")
POSITIONS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_positions.csv")
COOLDOWN_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_cooldown.csv")
HIGH_LOW_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_high_low.csv")
PAIR_STATUS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_pair_status.csv")
BUY_QUEUE_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_buy_queue.csv")
FUNDS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_funds.csv")
BLACKLIST_FILE = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_blacklist.csv")
OBSERVE_LIST_FILE = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_observe_list.csv")

TRADE_HEADER = ["timestamp", "inst_id", "direction", "price", "qty", "amount", "remaining_funds", "position_price", "profit", "profit_pct", "reason"]
POSITIONS_HEADER = ["inst_id", "position_price", "position_qty", "position_amount", "buy_time", "buy_reason", "pair_type", "pair_sell_line", "peak_price", "stop_price", "profit_triggered", "check_level", "last_check_time", "current_price", "market_value", "unrealized_pnl", "unrealized_pnl_pct"]
COOLDOWN_HEADER = ["inst_id", "sell_time", "cooldown_until"]
HIGH_LOW_HEADER = ["inst_id", "all_time_high", "all_time_low", "last_update"]
PAIR_STATUS_HEADER = ["inst_id", "pair_type", "is_paired", "buy_price", "buy_time", "sell_line", "sell_line_price", "is_sold", "sell_time", "sell_reason", "profit", "profit_pct"]
BUY_QUEUE_HEADER = ["inst_id", "status", "ref_high", "ref_low", "add_time", "last_check_time"]
FUNDS_HEADER = ["remaining_funds", "total_position_market_value", "total_unrealized_pnl", "total_realized_pnl", "net_asset_value", "last_update"]
BLACKLIST_HEADER = ["inst_id", "reason", "add_time"]
OBSERVE_LIST_HEADER = ["inst_id", "reason", "add_time"]

CSV_LOCK = threading.RLock()

def _ensure_csv_header(file_path: str, header: List[str]):
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        try:
            with CSV_LOCK, open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                csv.writer(f).writerow(header)
        except Exception as e:
            logging.error(f"初始化CSV失败 {file_path}: {e}")

def _read_csv(file_path: str, header: List[str]) -> List[Dict]:
    _ensure_csv_header(file_path, header)
    result = []
    try:
        with CSV_LOCK, open(file_path, 'r', encoding='utf-8-sig') as f:
            for row in csv.DictReader(f):
                result.append(row)
    except Exception:
        pass
    return result

def _write_csv_row(file_path: str, header: List[str], row: Dict):
    try:
        _ensure_csv_header(file_path, header)
        with CSV_LOCK, open(file_path, 'a', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerow([row.get(c, "") for c in header])
    except Exception as e:
        logging.error(f"写入CSV失败 {file_path}: {e}")

def _write_csv_all(file_path: str, header: List[str], rows: List[Dict]):
    try:
        with CSV_LOCK, open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=header)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
    except Exception as e:
        logging.error(f"写入CSV失败 {file_path}: {e}")

def load_blacklist() -> set:
    rows = _read_csv(BLACKLIST_FILE, BLACKLIST_HEADER)
    return {row.get("inst_id", "") for row in rows if row.get("inst_id")}

def add_to_blacklist(inst_id: str, reason: str):
    if inst_id in load_blacklist():
        return
    _write_csv_row(BLACKLIST_FILE, BLACKLIST_HEADER,
                   {"inst_id": inst_id, "reason": reason, "add_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    logging.info(f"[黑名单] {inst_id} 已加入，原因: {reason}")

def load_observe_list() -> set:
    rows = _read_csv(OBSERVE_LIST_FILE, OBSERVE_LIST_HEADER)
    return {row.get("inst_id", "") for row in rows if row.get("inst_id")}

def add_to_observe_list(inst_id: str, reason: str):
    if inst_id in load_observe_list():
        return
    _write_csv_row(OBSERVE_LIST_FILE, OBSERVE_LIST_HEADER,
                   {"inst_id": inst_id, "reason": reason, "add_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    logging.info(f"[待观察] {inst_id} 已加入，原因: {reason}")

def load_positions() -> Dict[str, Dict]:
    rows = _read_csv(POSITIONS_FILE, POSITIONS_HEADER)
    result = {}
    for row in rows:
        inst_id = row.get("inst_id", "")
        if inst_id:
            result[inst_id] = {
                "position_price": float(row.get("position_price", 0) or 0),
                "position_qty": float(row.get("position_qty", 0) or 0),
                "position_amount": float(row.get("position_amount", 0) or 0),
                "buy_time": row.get("buy_time", ""),
                "buy_reason": row.get("buy_reason", ""),
                "pair_type": row.get("pair_type", ""),
                "pair_sell_line": row.get("pair_sell_line", ""),
                "peak_price": float(row.get("peak_price", 0) or 0),
                "stop_price": float(row.get("stop_price", 0) or 0),
                "profit_triggered": (str(row.get("profit_triggered") or "False")).strip().lower() == "true",
                "check_level": row.get("check_level", "T3"),
                "last_check_time": row.get("last_check_time", ""),
                "current_price": float(row.get("current_price", 0) or 0),
                "market_value": float(row.get("market_value", 0) or 0),
                "unrealized_pnl": float(row.get("unrealized_pnl", 0) or 0),
                "unrealized_pnl_pct": float(row.get("unrealized_pnl_pct", 0) or 0),
            }
    return result

def save_positions(positions: Dict[str, Dict]):
    rows = []
    for inst_id, data in positions.items():
        rows.append({
            "inst_id": inst_id,
            "position_price": data.get("position_price", 0),
            "position_qty": data.get("position_qty", 0),
            "position_amount": data.get("position_amount", 0),
            "buy_time": data.get("buy_time", ""),
            "buy_reason": data.get("buy_reason", ""),
            "pair_type": data.get("pair_type", ""),
            "pair_sell_line": data.get("pair_sell_line", ""),
            "peak_price": data.get("peak_price", 0),
            "stop_price": data.get("stop_price", 0),
            "profit_triggered": "True" if data.get("profit_triggered", False) else "False",
            "check_level": data.get("check_level", "T3"),
            "last_check_time": data.get("last_check_time", ""),
            "current_price": data.get("current_price", 0),
            "market_value": data.get("market_value", 0),
            "unrealized_pnl": data.get("unrealized_pnl", 0),
            "unrealized_pnl_pct": data.get("unrealized_pnl_pct", 0),
        })
    _write_csv_all(POSITIONS_FILE, POSITIONS_HEADER, rows)

def load_cooldown() -> Dict[str, str]:
    rows = _read_csv(COOLDOWN_FILE, COOLDOWN_HEADER)
    return {row["inst_id"]: row["cooldown_until"] for row in rows if row.get("inst_id")}

def save_cooldown_entry(inst_id: str, cooldown_until: str):
    _write_csv_row(COOLDOWN_FILE, COOLDOWN_HEADER,
                   {"inst_id": inst_id, "sell_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "cooldown_until": cooldown_until})

def load_high_low() -> Dict[str, Dict]:
    rows = _read_csv(HIGH_LOW_FILE, HIGH_LOW_HEADER)
    res = {}
    for r in rows:
        res[r["inst_id"]] = {
            "all_time_high": float(r.get("all_time_high", 0) or 0),
            "all_time_low": float(r.get("all_time_low", 0) or 0),
            "last_update": r.get("last_update", "")
        }
    return res

def save_high_low_entry(inst_id: str, high: float, low: float):
    rows = _read_csv(HIGH_LOW_FILE, HIGH_LOW_HEADER)
    found = False
    for row in rows:
        if row["inst_id"] == inst_id:
            row["all_time_high"] = str(high)
            row["all_time_low"] = str(low)
            row["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            found = True
            break
    if not found:
        rows.append({"inst_id": inst_id, "all_time_high": str(high), "all_time_low": str(low), "last_update": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    _write_csv_all(HIGH_LOW_FILE, HIGH_LOW_HEADER, rows)

def load_pair_status() -> Dict[str, Dict]:
    rows = _read_csv(PAIR_STATUS_FILE, PAIR_STATUS_HEADER)
    res = {}
    for r in rows:
        res[r["inst_id"]] = {
            "pair_type": r.get("pair_type", ""),
            "is_paired": str(r.get("is_paired", "False")).lower() == "true",
            "buy_price": float(r.get("buy_price", 0) or 0),
            "buy_time": r.get("buy_time", ""),
            "sell_line": r.get("sell_line", ""),
            "sell_line_price": float(r.get("sell_line_price", 0) or 0),
            "is_sold": str(r.get("is_sold", "False")).lower() == "true",
            "sell_time": r.get("sell_time", ""),
            "sell_reason": r.get("sell_reason", ""),
            "profit": float(r.get("profit", 0) or 0),
            "profit_pct": float(r.get("profit_pct", 0) or 0),
        }
    return res

def save_pair_status(inst_id: str, data: Dict):
    rows = _read_csv(PAIR_STATUS_FILE, PAIR_STATUS_HEADER)
    found = False
    for row in rows:
        if row["inst_id"] == inst_id:
            row.update({k: str(v) for k, v in data.items() if k in row})
            found = True
            break
    if not found:
        rows.append({
            "inst_id": inst_id, "pair_type": data.get("pair_type", ""),
            "is_paired": "True" if data.get("is_paired") else "False",
            "buy_price": str(data.get("buy_price",0)),
            "buy_time": data.get("buy_time",""),
            "sell_line": data.get("sell_line",""),
            "sell_line_price": str(data.get("sell_line_price",0)),
            "is_sold": "True" if data.get("is_sold") else "False",
            "sell_time": data.get("sell_time",""),
            "sell_reason": data.get("sell_reason",""),
            "profit": str(data.get("profit",0)),
            "profit_pct": str(data.get("profit_pct",0)),
        })
    _write_csv_all(PAIR_STATUS_FILE, PAIR_STATUS_HEADER, rows)

def load_buy_queue() -> List[Dict]:
    return _read_csv(BUY_QUEUE_FILE, BUY_QUEUE_HEADER)

def save_buy_queue(rows: List[Dict]):
    _write_csv_all(BUY_QUEUE_FILE, BUY_QUEUE_HEADER, rows)

def save_trade(row: Dict):
    _write_csv_row(TRADE_FILE, TRADE_HEADER, row)

# ==================== 止盈表（完整） ====================
PROFIT_STOP_MAP = [
    (0.065, 0.015), (0.066, 0.0149), (0.067, 0.0148), (0.068, 0.0147),
    (0.069, 0.0146), (0.070, 0.0145), (0.071, 0.0144), (0.072, 0.0143),
    (0.073, 0.0142), (0.074, 0.0141), (0.075, 0.0140), (0.076, 0.0139),
    (0.077, 0.0138), (0.078, 0.0137), (0.079, 0.0136), (0.080, 0.0135),
    (0.081, 0.0134), (0.082, 0.0133), (0.083, 0.0132), (0.084, 0.0131),
    (0.085, 0.0130), (0.086, 0.0129), (0.087, 0.0128), (0.088, 0.0127),
    (0.089, 0.0126), (0.090, 0.0125), (0.091, 0.0124), (0.092, 0.0123),
    (0.093, 0.0122), (0.094, 0.0121), (0.095, 0.0120), (0.096, 0.0119),
    (0.097, 0.0118), (0.098, 0.0117), (0.099, 0.0116), (0.100, 0.0115),
    (0.101, 0.0114), (0.102, 0.0113), (0.103, 0.0112), (0.104, 0.0111),
    (0.105, 0.0110), (0.106, 0.0109), (0.107, 0.0108), (0.108, 0.0107),
    (0.109, 0.0106), (0.110, 0.0105), (0.111, 0.0104), (0.112, 0.0103),
    (0.113, 0.0102), (0.114, 0.0101), (0.115, 0.0100), (0.116, 0.0099),
    (0.117, 0.0098), (0.118, 0.0097), (0.119, 0.0096), (0.120, 0.0095),
    (0.121, 0.0094), (0.122, 0.0093), (0.123, 0.0092), (0.124, 0.0091),
    (0.125, 0.0090), (0.126, 0.0089), (0.127, 0.0088), (0.128, 0.0087),
    (0.129, 0.0086), (0.130, 0.0085), (0.131, 0.0084), (0.132, 0.0083),
    (0.133, 0.0082), (0.134, 0.0081), (0.135, 0.0080), (0.136, 0.0079),
    (0.137, 0.0078), (0.138, 0.0077), (0.139, 0.0076), (0.140, 0.0075),
    (0.141, 0.0074), (0.142, 0.0073), (0.143, 0.0072), (0.144, 0.0071),
    (0.145, 0.0070), (0.146, 0.0069), (0.147, 0.0068), (0.148, 0.0067),
    (0.149, 0.0066), (0.150, 0.0065), (0.151, 0.0064), (0.152, 0.0063),
    (0.153, 0.0062), (0.154, 0.0061), (0.155, 0.0060), (0.156, 0.0059),
    (0.157, 0.0058), (0.158, 0.0057), (0.159, 0.0056), (0.160, 0.0055),
    (0.161, 0.0054), (0.162, 0.0053), (0.163, 0.0052), (0.164, 0.0051),
    (0.165, 0.0050), (0.166, 0.0049), (0.167, 0.0048), (0.168, 0.0047),
    (0.169, 0.0046), (0.170, 0.0045), (0.171, 0.0044), (0.172, 0.0043),
    (0.173, 0.0042), (0.174, 0.0041), (0.175, 0.0040), (0.176, 0.0039),
    (0.177, 0.0038), (0.178, 0.0037), (0.179, 0.0036), (0.180, 0.0035),
    (0.181, 0.0034), (0.182, 0.0033), (0.183, 0.0032), (0.184, 0.0031),
    (0.185, 0.0030), (0.186, 0.0029), (0.187, 0.0028), (0.188, 0.0027),
    (0.189, 0.0026), (0.190, 0.0025), (0.191, 0.0024), (0.192, 0.0023),
    (0.193, 0.0022), (0.194, 0.0021), (0.195, 0.0020), (0.196, 0.0019),
    (0.197, 0.0018), (0.198, 0.0017), (0.199, 0.0016), (0.200, 0.0015),
    (0.201, 0.0014), (0.202, 0.0013), (0.203, 0.0012), (0.204, 0.0011),
    (0.205, 0.0010), (0.206, 0.0009), (0.207, 0.0008), (0.208, 0.0007),
    (0.209, 0.0006), (0.210, 0.0005), (0.211, 0.0004), (0.212, 0.0003),
    (0.213, 0.0002), (0.214, 0.0001), (0.215, 0.0), (0.216, 0),
]

def get_stop_loss(profit_pct: float) -> float:
    for p, s in reversed(PROFIT_STOP_MAP):
        if profit_pct >= p:
            return s
    return 0.015

# ==================== PriceLevels ====================
class PriceLevels:
    def __init__(self, high: float, low: float):
        self.high = high
        self.low = low
        self.update(high, low)
    def update(self, high: float, low: float):
        self.high = high
        self.low = low
        self.high_mid = (high + low) / 2 if high != low else high
        self.pullback1_mid = (high + self.high_mid) / 2
        self.pullback_up = (high + self.pullback1_mid) / 2
        self.pullback2_down = (self.pullback1_mid + self.high_mid) / 2
        self.pullback3 = (self.pullback2_down + self.high_mid) / 2
        self.pullback4 = (self.pullback3 + self.high_mid) / 2
        self.mid = (high + low) / 2
        self.callback1_mid = (self.high_mid + self.mid) / 2
        self.callback_up = (self.high_mid + self.callback1_mid) / 2
        self.callback2_down = (self.callback1_mid + self.mid) / 2
        self.callback3 = (self.callback2_down + self.mid) / 2
        self.callback4 = (self.callback3 + self.mid) / 2
        self.low_mid = (self.mid + low) / 2
        self.transition1_mid = (self.mid + self.low_mid) / 2
        self.transition_up = (self.mid + self.transition1_mid) / 2
        self.transition2_down = (self.transition1_mid + self.low_mid) / 2
        self.transition3 = (self.transition2_down + self.low_mid) / 2
        self.transition4 = (self.transition3 + self.low_mid) / 2
        self.limit1_mid = (self.low_mid + low) / 2
        self.limit_up = (self.low_mid + self.limit1_mid) / 2
        self.limit2_down = (self.limit1_mid + low) / 2
        self.limit3 = (self.limit2_down + low) / 2
        self.limit4 = (self.limit3 + low) / 2
        self.buy_levels = {
            "pullback2_down": self.pullback2_down,
            "callback2_down": self.callback2_down,
            "transition2_down": self.transition2_down,
            "limit2_down": self.limit2_down,
        }
        self.sell_levels = {
            "pullback_up": self.pullback_up,
            "callback_up": self.callback_up,
            "transition_up": self.transition_up,
            "limit_up": self.limit_up,
        }
        self.pair_map = {
            "pullback2_down": ("pullback_up", "回撤"),
            "callback2_down": ("callback_up", "回调"),
            "transition2_down": ("transition_up", "过渡"),
            "limit2_down": ("limit_up", "极限"),
        }

# ==================== FutuClient（修正版） ====================
class FutuTradeError(Exception):
    pass

class FutuClient:
    def __init__(self, market: str, trade_mode: str):
        self.market = market
        self.trade_mode = trade_mode
        self.host = os.getenv("FUTU_OPEND_HOST", "127.0.0.1")
        self.port = _env_int("FUTU_OPEND_PORT", 11111)
        self.acc_id = None
        self.trd_env = None
        self.quote_ctx = None
        self.trade_ctx = None
        self.lot_cache = {}
        self._connect()

    def _connect(self):
        temp_ctx = ft.OpenSecTradeContext(host=self.host, port=self.port)
        ret, acc_df = temp_ctx.get_acc_list()
        temp_ctx.close()
        if ret != ft.RET_OK:
            raise FutuTradeError(f"获取账户列表失败: {acc_df}")

        if self.trade_mode == "demo":
            self.trd_env = ft.TrdEnv.SIMULATE
        else:
            if self.market == "CN":
                raise FutuTradeError("A股不支持实盘交易")
            self.trd_env = ft.TrdEnv.REAL

        filtered = acc_df[acc_df["trd_env"] == self.trd_env]
        if filtered.empty:
            env_name = "模拟盘" if self.trd_env == ft.TrdEnv.SIMULATE else "实盘"
            raise FutuTradeError(f"未找到{env_name}账户")

        acc_id_env = os.getenv("FUTU_ACCOUNT_ID")
        if acc_id_env:
            try:
                acc_id_env = int(acc_id_env)
                filtered = filtered[filtered["acc_id"] == acc_id_env]
                if filtered.empty:
                    raise FutuTradeError(f"指定账户ID {acc_id_env} 不存在")
            except ValueError:
                raise FutuTradeError(f"FUTU_ACCOUNT_ID 必须为整数")

        self.acc_id = filtered.iloc[0]["acc_id"]
        logging.info(f"使用账户ID: {self.acc_id} ({'模拟' if self.trd_env == ft.TrdEnv.SIMULATE else '实盘'})")

        self.quote_ctx = ft.OpenQuoteContext(host=self.host, port=self.port)
        # 关键：OpenSecTradeContext 不支持 acc_id，只传 host 和 port
        self.trade_ctx = ft.OpenSecTradeContext(host=self.host, port=self.port)

    def close(self):
        if self.quote_ctx:
            self.quote_ctx.close()
        if self.trade_ctx:
            self.trade_ctx.close()

    def get_acc_balance(self) -> float:
        ret, data = self.trade_ctx.accinfo_query(trd_env=self.trd_env, acc_id=self.acc_id)
        if ret != ft.RET_OK:
            raise FutuTradeError(f"获取账户信息失败: {data}")
        col = "avail_balance" if "avail_balance" in data.columns else "avail_withdraw_cash"
        return float(data.loc[0, col])

    def get_positions(self) -> Dict[str, Dict]:
        ret, data = self.trade_ctx.position_list_query(trd_env=self.trd_env, acc_id=self.acc_id)
        if ret != ft.RET_OK:
            raise FutuTradeError(f"获取持仓失败: {data}")
        pos_dict = {}
        for _, row in data.iterrows():
            code = row["code"]
            qty = float(row["qty"])
            if qty > 0:
                pos_dict[code] = {
                    "position_price": float(row["cost_price"]),
                    "position_qty": qty,
                    "position_amount": float(row["cost_price"]) * qty
                }
        return pos_dict

    def _get_lot_size(self, code: str) -> int:
        if code in self.lot_cache:
            return self.lot_cache[code]
        try:
            ret, data = self.quote_ctx.get_stock_basicinfo([code])
            if ret == ft.RET_OK and not data.empty:
                lot = int(data.iloc[0]["lot_size"])
                self.lot_cache[code] = lot
                return lot
        except Exception:
            pass
        if code.startswith("HK.") or code.startswith("SH.") or code.startswith("SZ."):
            lot = 100
        else:
            lot = 1
        self.lot_cache[code] = lot
        return lot

    def market_buy_amount(self, code: str, amount: float, ref_price: float = 0.0) -> Dict:
        if self.market == "CN" and self.trd_env == ft.TrdEnv.REAL:
            return {"error": "A股禁止实盘下单"}

        price = ref_price if ref_price > 0 else self._get_last_price(code)
        if price <= 0:
            return {"error": f"无法获取参考价: {code}"}

        lot_size = self._get_lot_size(code)
        qty = int(amount / price)
        qty = max(lot_size, (qty // lot_size) * lot_size)
        if qty <= 0:
            return {"error": f"折算数量为0"}

        ret, ret_data = self.trade_ctx.place_order(
            price=0, qty=qty, code=code,
            order_type=ft.OrderType.MARKET,
            trd_side=ft.TrdSide.BUY,
            trd_env=self.trd_env,
            acc_id=self.acc_id
        )
        if ret != ft.RET_OK:
            return {"error": f"下单失败:{ret_data}"}
        return {"price": price, "qty": qty, "cost": amount, "ord_id": self._extract_order_id(ret_data), "error": None}

    def market_sell_qty(self, code: str, qty: float) -> Dict:
        if self.market == "CN" and self.trd_env == ft.TrdEnv.REAL:
            return {"error": "A股禁止实盘下单"}
        qty_int = int(qty)
        if qty_int <= 0:
            return {"error": f"卖出数量无效"}
        ret, ret_data = self.trade_ctx.place_order(
            price=0, qty=qty_int, code=code,
            order_type=ft.OrderType.MARKET,
            trd_side=ft.TrdSide.SELL,
            trd_env=self.trd_env,
            acc_id=self.acc_id
        )
        if ret != ft.RET_OK:
            return {"error": f"卖出失败:{ret_data}"}
        return {"price": 0, "qty": qty_int, "proceeds": 0, "ord_id": self._extract_order_id(ret_data), "error": None}

    def _extract_order_id(self, ret_data) -> str:
        try:
            return str(ret_data["order_id"].iloc[0])
        except Exception:
            return str(ret_data)

    def _get_last_price(self, code: str) -> float:
        try:
            ret, data = self.quote_ctx.get_market_snapshot([code])
            if ret == ft.RET_OK and data is not None and len(data) > 0:
                return float(data.iloc[0]["last_price"])
        except Exception:
            pass
        return 0.0

    def subscribe_tick(self, code_list: List[str]):
        ret, err = self.quote_ctx.subscribe(code_list, [ft.SubType.TICKER])
        if ret != ft.RET_OK:
            raise FutuTradeError(f"订阅tick失败:{err}")

# ==================== SymbolState（业务逻辑） ====================
class SymbolState:
    def __init__(self, inst_id, ref_high, ref_low, system_params=None):
        self.inst_id = inst_id
        self.params = system_params or {}
        self.min_buy_amount = _env_float("MIN_BUY_AMOUNT", 100.0)
        self.client = _futu_client
        self.blacklist = load_blacklist()
        self.observe_list = load_observe_list()
        self._first_fail_time = 0.0
        self._fail_count_in_window = 0
        self._retry_after = 0.0
        self._window_duration = 600
        self._retry_delay = 120
        self.buy_signal_activated = False
        self.buy_signal_lowest_price = 0.0
        self.buy_signal_buy_level = ""
        self.buy_signal_buy_price = 0.0
        self.buy_signal_attempt_time = 0.0

        self.positions = load_positions()
        self.cooldown = load_cooldown()
        self.high_low = load_high_low()
        self.pair_status = load_pair_status()
        self.buy_queue = load_buy_queue()

        hl = self.high_low.get(inst_id, {})
        self.high = hl.get("all_time_high", ref_high)
        self.low = hl.get("all_time_low", ref_low)
        self.levels = PriceLevels(self.high, self.low)

        pos = self.positions.get(inst_id, {})
        self.has_position = pos.get("position_qty", 0) > 0
        self.position_price = pos.get("position_price", 0)
        self.position_qty = pos.get("position_qty", 0)
        self.position_amount = pos.get("position_amount", 0)
        self.buy_time = pos.get("buy_time", "")
        self.buy_reason = pos.get("buy_reason", "")
        self.pair_type = pos.get("pair_type", "")
        self.pair_sell_line = pos.get("pair_sell_line", "")
        self.peak_price = pos.get("peak_price", 0)
        self.stop_price = pos.get("stop_price", 0)
        self.profit_triggered = pos.get("profit_triggered", False)
        self.check_level = pos.get("check_level", "T3")
        self.last_check_time = pos.get("last_check_time", "")

        pair = self.pair_status.get(inst_id, {})
        self.is_paired = pair.get("is_paired", False)
        self.pair_buy_price = pair.get("buy_price", 0)
        self.pair_buy_time = pair.get("buy_time", "")
        self.pair_sell_line_price = pair.get("sell_line_price", 0)
        self.is_sold = pair.get("is_sold", False)
        self.sell_time = pair.get("sell_time", "")
        self.sell_reason = pair.get("sell_reason", "")

        self._last_profit = 0.0
        self._last_profit_pct = 0.0
        self._balance_insufficient_until = 0.0
        self._sync_buy_queue()
        self._last_processed_price = 0.0
        self._last_processed_ts = 0.0

        if self.client is not None:
            refresh_positions_cache(force=True)
            if has_api_position(inst_id):
                if not self.has_position:
                    logging.info(f"{inst_id} API检测持仓，本地同步更新")
                    api_pos = get_api_positions().get(inst_id, {})
                    self.has_position = True
                    self.position_qty = api_pos.get("position_qty", 0)
                    self.position_price = api_pos.get("position_price", 0)
                    self.position_amount = api_pos.get("position_amount", 0)
                    self._save_state()

    def _sync_buy_queue(self):
        in_queue = False
        changed = False
        for row in self.buy_queue:
            if row.get("inst_id") == self.inst_id:
                in_queue = True
                if self.has_position:
                    new_status = "已持仓_排除"
                elif self._is_in_cooldown():
                    new_status = "静默期_排除"
                elif self.inst_id in self.observe_list:
                    new_status = "观察列表_排除"
                elif self.inst_id in self.blacklist:
                    new_status = "黑名单_排除"
                else:
                    new_status = "待买入"
                if row.get("status") != new_status:
                    row["status"] = new_status
                    row["last_check_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    changed = True
                break
        if not in_queue and not self.has_position and not self._is_in_cooldown() and self.inst_id not in self.observe_list and self.inst_id not in self.blacklist:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.buy_queue.append({
                "inst_id": self.inst_id,
                "status": "待买入",
                "ref_high": str(self.high),
                "ref_low": str(self.low),
                "add_time": now_str,
                "last_check_time": now_str
            })
            changed = True
        if changed:
            save_buy_queue(self.buy_queue)

    def _is_in_cooldown(self):
        until = self.cooldown.get(self.inst_id, "")
        if until:
            try:
                return datetime.now() < datetime.strptime(until, "%Y-%m-%d %H:%M:%S")
            except Exception:
                pass
        return False

    def _update_check_level(self, profit_pct):
        if profit_pct >= 0.05:
            self.check_level = "T0"
        elif profit_pct > 0:
            self.check_level = "T1"
        elif self.pair_type in ("回撤", "回调", "过渡", "极限"):
            self.check_level = "T2"
        else:
            self.check_level = "T3"
        self.last_check_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _can_buy(self):
        if self.has_position:
            return False, "本地已有持仓"
        if has_api_position(self.inst_id):
            return False, "API检测到现有持仓"
        if self._is_in_cooldown():
            return False, f"静默期中 (至 {self.cooldown.get(self.inst_id, '')})"
        if self.inst_id in self.observe_list:
            return False, "在待观察列表中"
        if self.inst_id in self.blacklist:
            return False, "在黑名单中"
        if time.time() < self._balance_insufficient_until:
            return False, f"余额不足冷却中（剩余 {int(self._balance_insufficient_until - time.time())} 秒）"
        if time.time() < self._retry_after:
            return False, f"未知错误重试延迟中（剩余 {int(self._retry_after - time.time())} 秒）"
        balance = get_futu_balance()
        if balance < self.min_buy_amount:
            self._balance_insufficient_until = time.time() + 300
            return False, f"资金不足: 余额 {balance:.2f} < 最小买入 {self.min_buy_amount:.2f}（冷却300秒）"
        raw_amount = balance * self.params.get("buy_ratio", 0.01)
        buy_amount = min(max(raw_amount, self.min_buy_amount), balance)
        if buy_amount <= 0:
            return False, "计算出的买入金额为0"
        if self.is_paired and not self.is_sold:
            return False, f"已配对未卖出 ({self.pair_type})"
        return True, ""

    def _can_sell(self):
        return (True, "") if self.has_position else (False, "无持仓")

    def _get_buy_level(self, price):
        buy_keys = ["pullback2_down", "callback2_down", "transition2_down", "limit2_down"]
        best_key = None
        best_price = None
        for key in buy_keys:
            level = getattr(self.levels, key)
            if price < level:
                if best_price is None or level < best_price:
                    best_price = level
                    best_key = key
        return best_key, best_price

    def _execute_buy(self, price, reason, buy_key, sell_key, pair_name):
        balance = get_futu_balance(force_refresh=True)
        if balance < self.min_buy_amount:
            self._balance_insufficient_until = time.time() + 300
            logging.info(f"{self.inst_id} 账户余额 {balance:.2f} 低于最小买入金额 {self.min_buy_amount:.2f}，跳过买入并冷却300秒")
            return None
        buy_level = getattr(self.levels, buy_key)
        if price >= buy_level:
            logging.debug(f"{self.inst_id} 当前价 {price:.8g} 不小于买点 {buy_level:.8g}，跳过")
            return None
        can, msg = self._can_buy()
        if not can:
            logging.debug(f"{self.inst_id} 买入跳过: {msg}")
            return None
        raw_amount = balance * self.params.get("buy_ratio", 0.01)
        buy_amount = max(raw_amount, self.min_buy_amount)
        if buy_amount > balance:
            buy_amount = balance
        fill = self.client.market_buy_amount(self.inst_id, buy_amount, ref_price=price)
        now = time.time()
        if fill is None:
            logging.error(f"{self.inst_id} API买入失败返回None")
            return self._handle_unknown_error(reason="未知错误返回None")
        if isinstance(fill, dict) and fill.get("error"):
            error_msg = fill["error"]
            logging.error(f"{self.inst_id} API买入失败:{error_msg}")
            return self._handle_unknown_error(reason=error_msg[:100])
        # 成交成功
        price = fill["price"] if fill["price"] != 0 else price
        qty = fill["qty"] if fill["qty"] != 0 else (buy_amount / price if price > 0 else 0)
        if qty <= 0:
            logging.error(f"{self.inst_id} 买入成交数量异常 qty={qty}，忽略该笔")
            return None
        buy_amount = fill["cost"]
        reason = f"{reason} | ordId={fill.get('ord_id','')}"
        logging.info(f"{self.inst_id} 买入成交 ordId={fill.get('ord_id')} 均价 {price:.8g} 数量 {qty:.8g} 花费 {buy_amount:.2f}")
        self._first_fail_time = 0.0
        self._fail_count_in_window = 0
        self._retry_after = 0.0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.has_position = True
        self.position_price = price
        self.position_qty = qty
        self.position_amount = buy_amount
        self.buy_time = now_str
        self.buy_reason = reason
        self.pair_type = pair_name
        self.pair_sell_line = sell_key
        self.is_paired = True
        self.pair_buy_price = price
        self.pair_buy_time = now_str
        self.is_sold = False
        self.sell_time = ""
        self.sell_reason = ""
        self.check_level = "T3"
        self.last_check_time = now_str
        self.pair_sell_line_price = getattr(self.levels, sell_key)
        self.profit_triggered = False
        self.peak_price = price
        self.stop_price = 0
        self._last_processed_price = price
        self.buy_signal_activated = False
        self.buy_signal_lowest_price = 0.0
        self.buy_signal_buy_level = ""
        self.buy_signal_buy_price = 0.0
        _update_funds_for_view(get_futu_balance(force_refresh=True))
        refresh_positions_cache(force=True)
        self._remove_from_buy_queue()
        self._save_state()
        save_trade({
            "timestamp": now_str,
            "inst_id": self.inst_id,
            "direction": "BUY",
            "price": price,
            "qty": qty,
            "amount": buy_amount,
            "remaining_funds": f"{get_futu_balance():.2f}",
            "position_price": price,
            "profit": "0",
            "profit_pct": "0",
            "reason": reason
        })
        msg = f"{self.inst_id} {reason} @ {price:.8g}, 金额 {buy_amount:.2f}, 买点: {buy_key}"
        logging.info(msg)
        push_to_dingtalk(os.getenv("DINGTALK_WEBHOOK_futu"), os.getenv("DINGTALK_SECRET_futu"),
                         f"💰 买入\n\n{self.inst_id}\n{pair_name}买点: {buy_key}\n价格: {price:.8g}\n金额: {buy_amount:.2f}")
        return msg

    def _handle_unknown_error(self, reason: str):
        now = time.time()
        if now - self._first_fail_time > self._window_duration:
            self._first_fail_time = now
            self._fail_count_in_window = 1
            self._retry_after = now + self._retry_delay
            logging.info(f"{self.inst_id} 未知错误第1次，{self._retry_delay}秒后重试")
            return None
        else:
            self._fail_count_in_window += 1
            if self._fail_count_in_window >= 2:
                logging.warning(f"{self.inst_id} 连续两次未知错误，加入待观察列表")
                add_to_observe_list(self.inst_id, f"未知错误连续失败:{reason[:100]}")
                self._remove_from_buy_queue()
                self._first_fail_time = 0.0
                self._fail_count_in_window = 0
                self._retry_after = 0.0
                return None
            else:
                self._retry_after = now + self._retry_delay
                logging.info(f"{self.inst_id} 未知错误第{self._fail_count_in_window}次，{self._retry_delay}秒后重试")
                return None

    def _remove_from_buy_queue(self):
        self.buy_queue = [r for r in self.buy_queue if r.get("inst_id") != self.inst_id]
        save_buy_queue(self.buy_queue)

    def _execute_sell(self, price, reason):
        can, msg = self._can_sell()
        if not can:
            logging.debug(f"{self.inst_id} 卖出跳过: {msg}")
            return None
        qty = self.position_qty
        fill = self.client.market_sell_qty(self.inst_id, qty)
        if not fill or fill.get("error"):
            logging.error(f"{self.inst_id} API卖出失败 {fill}")
            push_to_dingtalk(os.getenv("DINGTALK_WEBHOOK_futu"), os.getenv("DINGTALK_SECRET_futu"),
                             f"⚠️卖出失败\n{self.inst_id}\n{reason}\n数量:{qty:.8g}")
            return None
        price = fill["price"] if fill["price"] != 0 else price
        qty = fill["qty"]
        amount = fill["proceeds"] if fill["proceeds"] != 0 else price * qty
        reason = f"{reason} | ordId={fill.get('ord_id','')}"
        logging.info(f"{self.inst_id} 卖出成交 ordId={fill.get('ord_id')} 均价 {price:.8g} 数量 {qty:.8g} 到手 {amount:.2f}")
        profit = amount - self.position_amount
        profit_pct = profit / self.position_amount * 100 if self.position_amount > 0 else 0
        cost_price = self.position_price
        cached_pair_type = self.pair_type
        cached_pair_sell_line = self.pair_sell_line
        cached_pair_sell_line_price = self.pair_sell_line_price
        cached_buy_time = self.buy_time
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._last_profit = profit
        self._last_profit_pct = profit_pct
        self.has_position = False
        self.position_price = 0
        self.position_qty = 0
        self.position_amount = 0
        self.buy_time = ""
        self.buy_reason = ""
        self.is_paired = False
        self.is_sold = True
        self.sell_time = now_str
        self.sell_reason = reason
        self.profit_triggered = False
        self.peak_price = 0
        self.stop_price = 0
        self.check_level = ""
        self.last_check_time = ""
        cooldown_days = self.params.get("cooldown_days",7)
        cooldown_until = (datetime.now() + timedelta(days=cooldown_days)).strftime("%Y-%m-%d %H:%M:%S")
        save_cooldown_entry(self.inst_id, cooldown_until)
        self.cooldown[self.inst_id] = cooldown_until
        self._add_to_buy_queue("静默期_排除")
        self.pair_buy_price = cost_price
        self.pair_buy_time = cached_buy_time
        self.pair_sell_line_price = cached_pair_sell_line_price
        self.pair_type = cached_pair_type
        self.pair_sell_line = cached_pair_sell_line
        _update_funds_for_view(get_futu_balance(force_refresh=True))
        refresh_positions_cache(force=True)
        self._save_state()
        save_trade({
            "timestamp": now_str,
            "inst_id": self.inst_id,
            "direction": "SELL",
            "price": price,
            "qty": qty,
            "amount": amount,
            "remaining_funds": f"{get_futu_balance():.2f}",
            "position_price": cost_price,
            "profit": f"{profit:.12g}",
            "profit_pct": f"{profit_pct:.6f}",
            "reason": reason
        })
        msg = f"{self.inst_id} {reason} @ {price:.8g}, 数量 {qty:.8g}, 盈亏 {profit:.2f} ({profit_pct:.2f}%)"
        logging.info(msg)
        push_to_dingtalk(os.getenv("DINGTALK_WEBHOOK_futu"), os.getenv("DINGTALK_SECRET_futu"),
                         f"💸 卖出\n\n{self.inst_id}\n{reason}\n价格: {price:.8g}\n盈亏: {profit:.2f} ({profit_pct:.2f}%)")
        return msg

    def _add_to_buy_queue(self, status):
        for row in self.buy_queue:
            if row.get("inst_id") == self.inst_id:
                row["status"] = status
                row["last_check_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                save_buy_queue(self.buy_queue)
                return
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.buy_queue.append({
            "inst_id": self.inst_id,
            "status": status,
            "ref_high": str(self.levels.high),
            "ref_low": str(self.levels.low),
            "add_time": now_str,
            "last_check_time": now_str
        })
        save_buy_queue(self.buy_queue)

    def _check_take_profit(self, price):
        if not self.has_position or self.position_price <= 0:
            return None
        profit_pct = (price - self.position_price) / self.position_price
        if price > self.peak_price:
            self.peak_price = price
        stop_ratio = get_stop_loss(profit_pct)
        if profit_pct >= 0.20:
            logging.info(f"{self.inst_id} 盈利≥20%立即清仓")
            return self._execute_sell(price, f"盈利率 {profit_pct*100:.1f}% ≥20%，清仓")
        new_stop = self.peak_price * (1 - stop_ratio)
        if new_stop > self.stop_price:
            self.stop_price = new_stop
        if profit_pct >= 0.05 and not self.profit_triggered:
            self.profit_triggered = True
            logging.info(f"{self.inst_id} 盈利率 {profit_pct*100:.2f}% ≥5%，启动动态止盈，当前止损价 {self.stop_price:.8g}")
        if self.profit_triggered and price <= self.stop_price:
            logging.info(f"{self.inst_id} 触发止盈清仓: 当前价 {price:.8g} ≤ 止盈价 {self.stop_price:.8g}")
            return self._execute_sell(price, f"止盈清仓 (峰值 {self.peak_price:.8g}, 止损 {stop_ratio*100:.1f}%)")
        return None

    def update_price(self, price, ts):
        alerts = []
        if price == self._last_processed_price and ts == self._last_processed_ts:
            return alerts
        self._last_processed_price = price
        self._last_processed_ts = ts
        self._sync_buy_queue()

        if price > self.high:
            self.high = price
            self.levels.update(self.high, self.low)
            save_high_low_entry(self.inst_id, self.high, self.low)
            alerts.append(f"创新高 {price:.8g}")
        if price < self.low:
            self.low = price
            self.levels.update(self.high, self.low)
            save_high_low_entry(self.inst_id, self.high, self.low)
            alerts.append(f"创新低 {price:.8g}")

        if self.has_position:
            profit_pct = (price - self.position_price) / self.position_price if self.position_price > 0 else 0
            self._update_check_level(profit_pct)
            if self.check_level == "T0":
                res = self._check_take_profit(price)
                if res:
                    alerts.append(f"止盈清仓 @ {price:.8g}")
                    self._save_state()
                    return alerts
            if self.is_paired and not self.is_sold:
                sell_price = self.pair_sell_line_price
                if price >= sell_price:
                    res = self._execute_sell(price, f"{self.pair_type}卖点 {self.pair_sell_line} 触发")
                    if res:
                        alerts.append(f"{self.pair_type}卖点清仓 @ {price:.8g}")
                        self._save_state()
                        return alerts

        if not self.has_position and not self._is_in_cooldown() and self.inst_id not in self.observe_list and self.inst_id not in self.blacklist:
            buy_key, buy_level = self._get_buy_level(price)
            if buy_key is not None:
                if not self.buy_signal_activated:
                    self.buy_signal_activated = True
                    self.buy_signal_lowest_price = price
                    self.buy_signal_buy_level = buy_key
                    self.buy_signal_buy_price = buy_level
                    self.buy_signal_attempt_time = 0.0
                    logging.debug(f"{self.inst_id} 买入信号激活，买点 {buy_key} ({buy_level:.8g})，当前价 {price:.8g}")
                else:
                    if price < self.buy_signal_lowest_price:
                        self.buy_signal_lowest_price = price
                        logging.debug(f"{self.inst_id} 更新最低点 {price:.8g}")
            else:
                if self.buy_signal_activated:
                    if price >= self.buy_signal_buy_price:
                        logging.debug(f"{self.inst_id} 价格回升至买点之上，取消买入信号")
                        self.buy_signal_activated = False
                        self.buy_signal_lowest_price = 0.0
                        self.buy_signal_buy_level = ""
                        self.buy_signal_buy_price = 0.0

            if self.buy_signal_activated:
                now = time.time()
                if now - self.buy_signal_attempt_time >= 30:
                    rebound_price = self.buy_signal_lowest_price * 1.01
                    if price >= rebound_price and price < self.buy_signal_buy_price:
                        logging.debug(f"{self.inst_id} 满足反弹买入条件：最低 {self.buy_signal_lowest_price:.8g}, 反弹1%={rebound_price:.8g}, 当前价 {price:.8g}")
                        sell_key, pair_name = self.levels.pair_map[self.buy_signal_buy_level]
                        res = self._execute_buy(
                            price,
                            f"反弹1%买入 (最低 {self.buy_signal_lowest_price:.8g})",
                            self.buy_signal_buy_level,
                            sell_key,
                            pair_name
                        )
                        self.buy_signal_attempt_time = now
                        if res:
                            alerts.append(f"反弹买入 @ {price:.8g}, 买点: {self.buy_signal_buy_level}")
                            self._save_state()
                            return alerts
        self._save_state()
        return alerts

    def _save_state(self):
        positions = load_positions()
        if self.has_position:
            cp = self._last_processed_price if self._last_processed_price > 0 else self.position_price
            qty = self.position_qty
            cost = self.position_amount
            mv = cp * qty if cp > 0 else cost
            upnl = mv - cost if cost > 0 else 0
            upnl_pct = (upnl / cost * 100.0) if cost > 0 else 0
            positions[self.inst_id] = {
                "position_price": self.position_price,
                "position_qty": self.position_qty,
                "position_amount": self.position_amount,
                "buy_time": self.buy_time,
                "buy_reason": self.buy_reason,
                "pair_type": self.pair_type,
                "pair_sell_line": self.pair_sell_line,
                "peak_price": self.peak_price,
                "stop_price": self.stop_price,
                "profit_triggered": self.profit_triggered,
                "check_level": self.check_level,
                "last_check_time": self.last_check_time,
                "current_price": cp,
                "market_value": mv,
                "unrealized_pnl": upnl,
                "unrealized_pnl_pct": upnl_pct,
            }
        else:
            positions.pop(self.inst_id, None)
        save_positions(positions)
        if self.is_paired or self.is_sold:
            payload = {
                "pair_type": self.pair_type,
                "is_paired": self.is_paired,
                "buy_price": self.pair_buy_price,
                "buy_time": self.pair_buy_time,
                "sell_line": self.pair_sell_line,
                "sell_line_price": self.pair_sell_line_price,
                "is_sold": self.is_sold,
                "sell_time": self.sell_time,
                "sell_reason": self.sell_reason,
            }
            if self.is_sold:
                payload["profit"] = f"{self._last_profit:.12g}"
                payload["profit_pct"] = f"{self._last_profit_pct:.6f}"
            save_pair_status(self.inst_id, payload)
        else:
            rows = _read_csv(PAIR_STATUS_FILE, PAIR_STATUS_HEADER)
            rows = [r for r in rows if r.get("inst_id") != self.inst_id]
            _write_csv_all(PAIR_STATUS_FILE, PAIR_STATUS_HEADER, rows)

# ==================== 全局变量与缓存 ====================
_futu_client: Optional[FutuClient] = None
_balance_cache = {"value":0.0, "timestamp":0.0}
_BALANCE_CACHE_TTL = 3.0
_positions_cache = {"data": {}, "timestamp":0.0}
_POSITIONS_CACHE_TTL = 30

def refresh_positions_cache(force=False):
    global _positions_cache
    now = time.time()
    if not force and (now - _positions_cache["timestamp"] < _POSITIONS_CACHE_TTL):
        return
    if _futu_client is None:
        return
    try:
        api_pos = _futu_client.get_positions()
        _positions_cache["data"] = api_pos
        _positions_cache["timestamp"] = now
    except Exception as e:
        logging.error(f"刷新持仓缓存失败: {e}")

def get_api_positions() -> Dict[str, Dict]:
    refresh_positions_cache()
    return _positions_cache["data"]

def has_api_position(inst_id: str) -> bool:
    positions = get_api_positions()
    return inst_id in positions and positions[inst_id].get("position_qty", 0) > 0

def get_futu_balance(force_refresh=False) -> float:
    global _balance_cache
    now = time.time()
    if not force_refresh and (now - _balance_cache["timestamp"] < _BALANCE_CACHE_TTL):
        return _balance_cache["value"]
    if _futu_client is None:
        return 0.0
    try:
        bal = _futu_client.get_acc_balance()
        _balance_cache["value"] = bal
        _balance_cache["timestamp"] = now
        return bal
    except Exception as e:
        logging.error(f"获取余额失败: {e}")
        return _balance_cache["value"]

def _update_funds_for_view(balance: float):
    try:
        row = {
            "remaining_funds": f"{balance:.2f}",
            "total_position_market_value": "0",
            "total_unrealized_pnl": "0",
            "total_realized_pnl": "0",
            "net_asset_value": f"{balance:.2f}",
            "last_update": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        _ensure_csv_header(FUNDS_FILE, FUNDS_HEADER)
        with CSV_LOCK, open(FUNDS_FILE, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=FUNDS_HEADER)
            writer.writeheader()
            writer.writerow(row)
    except Exception as e:
        logging.error(f"更新资金查看文件失败: {e}")

# ==================== MarketEngine ====================
class MarketEngine:
    def __init__(self, system_params=None):
        self.params = system_params or {}
        self.states: Dict[str, SymbolState] = {}
        self.lock = threading.Lock()

    def add_state(self, inst_id, ref_high, ref_low):
        state = SymbolState(inst_id, ref_high, ref_low, self.params)
        self.states[inst_id] = state
        hl = load_high_low()
        if inst_id in hl:
            state.high = hl[inst_id]["all_time_high"]
            state.low = hl[inst_id]["all_time_low"]
            state.levels.update(state.high, state.low)
        pos = load_positions()
        if inst_id in pos:
            data = pos[inst_id]
            state.has_position = data.get("position_qty", 0) > 0
            state.position_price = data.get("position_price", 0)
            state.position_qty = data.get("position_qty", 0)
            state.position_amount = data.get("position_amount", 0)
            state.buy_time = data.get("buy_time", "")
            state.buy_reason = data.get("buy_reason", "")
            state.pair_type = data.get("pair_type", "")
            state.pair_sell_line = data.get("pair_sell_line", "")
            state.peak_price = data.get("peak_price", 0)
            state.stop_price = data.get("stop_price", 0)
            state.profit_triggered = data.get("profit_triggered", False)
            state.check_level = data.get("check_level", "T3")
            cp = data.get("current_price", 0)
            if cp and float(cp) > 0:
                state._last_processed_price = float(cp)
        pair = load_pair_status()
        if inst_id in pair:
            data = pair[inst_id]
            state.is_paired = data.get("is_paired", False)
            state.pair_buy_price = data.get("buy_price", 0)
            state.pair_buy_time = data.get("buy_time", "")
            state.pair_sell_line_price = data.get("sell_line_price", 0)
            state.is_sold = data.get("is_sold", False)
            state.sell_time = data.get("sell_time", "")
            state.sell_reason = data.get("sell_reason", "")
            if data.get("pair_type"):
                state.pair_type = data["pair_type"]
            state._last_profit = data.get("profit", 0.0)
            state._last_profit_pct = data.get("profit_pct", 0.0)
        cooldown = load_cooldown()
        if inst_id in cooldown:
            state.cooldown[inst_id] = cooldown[inst_id]
        state._sync_buy_queue()

    def on_tick(self, inst_id: str, price: float, ts: float):
        with self.lock:
            state = self.states.get(inst_id)
            if state is None:
                return
            alerts = state.update_price(price, ts)
        if alerts:
            key_alert = [a for a in alerts if any(k in a for k in ("止盈", "清仓", "买入", "卖出"))]
            if key_alert:
                logging.info(f"[{inst_id}] " + " | ".join(key_alert))
            else:
                logging.debug(f"[{inst_id}] " + " | ".join(alerts))

# tick回调
market_engine: Optional[MarketEngine] = None

def tick_callback(quote_ctx, ret_code, content):
    if ret_code != ft.RET_OK:
        return
    for item in content:
        inst_id = item["code"]
        price = float(item["price"])
        raw_time = item.get("time", "")
        if raw_time:
            try:
                ts = datetime.strptime(raw_time, "%Y-%m-%d %H:%M:%S.%f").timestamp()
            except ValueError:
                try:
                    ts = datetime.strptime(raw_time, "%Y-%m-%d %H:%M:%S").timestamp()
                except Exception:
                    ts = time.time()
        else:
            ts = time.time()
        if market_engine:
            market_engine.on_tick(inst_id, price, ts)

# ==================== 主程序 ====================
def main():
    global _futu_client, market_engine

    setup_logging()
    logging.info("=" * 60)
    logging.info(f"{PROGRAM_NAME} {VERSION}")
    logging.info(f"工作目录: {BASE_DIR}")
    logging.info(f"数据目录: {DATA_DIR}")

    trade_mode = os.getenv("TRADE_MODE", "demo").strip().lower()
    if trade_mode not in ("demo", "live"):
        trade_mode = "demo"
    market = os.getenv("MARKET", "HK").strip().upper()
    if market not in ("CN", "HK", "US"):
        market = "HK"

    logging.info(f"运行市场: {market} | 交易模式: {'模拟盘' if trade_mode == 'demo' else '实盘'}")

    # 加载配置
    try:
        config, system_params = load_config()
        logging.info(f"加载股票池成功: {len(config)} 个标的")
    except Exception as e:
        logging.critical(f"配置加载失败: {e}", exc_info=True)
        return
    if not config:
        logging.error("未找到任何启用的标的")
        return

    # 获取所有市场类型
    markets_set = set()
    for code in config.keys():
        if code.startswith('HK.'):
            markets_set.add('HK')
        elif code.startswith('US.'):
            markets_set.add('US')
        elif code.startswith('SH.') or code.startswith('SZ.'):
            markets_set.add('CN')
    if not markets_set:
        markets_set.add(market)
    logging.info(f"监控市场: {markets_set}")

    # 等待开盘
    wait_until_market_open(list(markets_set))

    # 初始化FutuClient
    try:
        _futu_client = FutuClient(market, trade_mode)
        bal = _futu_client.get_acc_balance()
        logging.info(f"FutuOpenD 连接成功，账户可用资金: {bal:.2f}")
        _update_funds_for_view(bal)
        refresh_positions_cache(force=True)
    except Exception as e:
        logging.critical(f"连接FutuOpenD失败: {e}", exc_info=True)
        return

    min_amt = _env_float("MIN_BUY_AMOUNT", 100.0)
    cooldown_days = _env_int("COOLDOWN_DAYS", 7)
    logging.info(f"单笔最小买入金额: {min_amt:.2f}")
    logging.info(f"交易静默期天数: {cooldown_days}")

    market_engine = MarketEngine(system_params)
    blacklist = load_blacklist()
    observe_list = load_observe_list()
    if blacklist:
        logging.info(f"黑名单数量: {len(blacklist)}")
    if observe_list:
        logging.info(f"待观察列表数量: {len(observe_list)}")

    code_list = list(config.keys())
    logging.info("初始化标的状态...")
    for inst_id, data in config.items():
        market_engine.add_state(inst_id, data["ref_high"], data["ref_low"])
        st = market_engine.states[inst_id]
        logging.debug(f"{inst_id}: 高{st.high:.6g} 低{st.low:.6g}")

    # 订阅行情
    try:
        _futu_client.subscribe_tick(code_list)
        _futu_client.quote_ctx.set_handler(tick_callback)
        logging.info(f"行情订阅完成，共 {len(code_list)} 个标的")
    except Exception as e:
        logging.critical(f"订阅tick行情失败: {e}", exc_info=True)
        return

    api_pos = get_api_positions()
    if api_pos:
        logging.info("==== 当前富途账户持仓 ====")
        for code, pos_data in api_pos.items():
            logging.info(f"  {code}: 数量 {pos_data['position_qty']:.6g}, 成本价 {pos_data['position_price']:.6g}")

    local_pos = load_positions()
    for inst, d in local_pos.items():
        logging.info(f"本地恢复持仓: {inst} qty={d['position_qty']:.6g} @ {d['position_price']:.6g}")

    buy_queue = load_buy_queue()
    pending = [x for x in buy_queue if x.get("status") == "待买入"]
    logging.info(f"待买入队列数量: {len(pending)}")
    logging.info(f"账户可用资金: {get_futu_balance(force_refresh=True):.2f}")

    logging.info("=" * 60)
    logging.info("策略运行中，Ctrl+C停止程序")
    logging.info("买入规则：跌破买点激活信号，从最低点反弹1%且仍低于买点执行买入")
    logging.info("=" * 60)

    last_check = time.time()
    try:
        while True:
            now = time.time()
            if now - last_check >= 60:
                if not any(is_market_open(m, datetime.utcnow()) for m in markets_set):
                    wait_until_market_open(list(markets_set))
                last_check = now
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("收到Ctrl+C，程序准备退出")
    finally:
        if _futu_client:
            _futu_client.close()
        _update_funds_for_view(get_futu_balance(force_refresh=True))
        logging.info("程序已退出")

# ==================== 无人托管入口 ====================
if __name__ == "__main__":
    while True:
        try:
            main()
            break
        except KeyboardInterrupt:
            logging.info("收到中断信号，退出")
            break
        except Exception as e:
            logging.critical(f"程序意外崩溃: {e}", exc_info=True)
            logging.info("等待 60 秒后自动重启...")
            time.sleep(60)