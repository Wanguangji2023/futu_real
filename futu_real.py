#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FutuOpenD A股/港股/美股自动交易
版本: v1.2
复刻原OKX逻辑：
1. 价格跌破任意买点（回撤2下/回调2下/过渡2下/极限2下）激活买入信号，记录最低点
2. 价格继续下跌更新最低点；价格自最低点反弹1%且仍低于买点执行买入
3. 买入成功或价格回升买点之上取消信号
4. A股仅模拟下单；港股/美股支持模拟/实盘；实盘模式.env开关控制
5. 黑名单、观察列表、静默冷却、动态移动止盈、钉钉告警、CSV持久化
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
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, List, Tuple
from dotenv import load_dotenv

import futu as ft

load_dotenv()

# ==================== 配置常量 ====================
PROGRAM_NAME = "futu_autoTrade"
VERSION = "v1.2"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_config.xlsx")
BLACKLIST_FILE = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_blacklist.csv")
OBSERVE_LIST_FILE = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_observe_list.csv")
DATA_DIR = os.path.join(BASE_DIR, f"{PROGRAM_NAME}_data")
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

TRADE_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_tradeRecord.csv")
POSITIONS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_positions.csv")
COOLDOWN_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_cooldown.csv")
HIGH_LOW_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_high_low.csv")
PAIR_STATUS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_pair_status.csv")
BUY_QUEUE_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_buy_queue.csv")
FUNDS_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_funds.csv")
LOG_FILE = os.path.join(DATA_DIR, f"{PROGRAM_NAME}_run.log")

TRADE_HEADER = ["timestamp", "inst_id", "direction", "price", "qty", "amount", "remaining_funds", "position_price", "profit", "profit_pct", "reason"]
POSITIONS_HEADER = ["inst_id", "position_price", "position_qty", "position_amount", "buy_time", "buy_reason", "pair_type", "pair_sell_line", "peak_price", "stop_price", "profit_triggered", "check_level", "last_check_time", "current_price", "market_value", "unrealized_pnl", "unrealized_pnl_pct"]
COOLDOWN_HEADER = ["inst_id", "sell_time", "cooldown_until"]
HIGH_LOW_HEADER = ["inst_id", "all_time_high", "all_time_low", "last_update"]
PAIR_STATUS_HEADER = ["inst_id", "pair_type", "is_paired", "buy_price", "buy_time", "sell_line", "sell_line_price", "is_sold", "sell_time", "sell_reason", "profit", "profit_pct"]
BUY_QUEUE_HEADER = ["inst_id", "status", "ref_high", "ref_low", "add_time", "last_check_time"]
FUNDS_HEADER = ["remaining_funds", "total_position_market_value", "total_unrealized_pnl", "total_realized_pnl", "net_asset_value", "last_update"]
BLACKLIST_HEADER = ["inst_id", "reason", "add_time"]
OBSERVE_LIST_HEADER = ["inst_id", "reason", "add_time"]

# FutuOpenD 连接配置
FUTU_OPEND_HOST = os.getenv("FUTU_OPEND_HOST", "127.0.0.1")
FUTU_OPEND_PORT = int(os.getenv("FUTU_OPEND_PORT", 11111))

# 市场类型：CN=A股，HK=港股，US=美股
RUN_MARKET = os.getenv("MARKET", "HK").strip().upper()
RAW_TRADE_MODE = os.getenv("TRADE_MODE", "demo").strip().lower()
# A股强制模拟，不允许实盘
if RUN_MARKET == "CN":
    TRADE_MODE = "demo"
else:
    TRADE_MODE = RAW_TRADE_MODE if RAW_TRADE_MODE in ("demo", "live") else "demo"

MODE_DESC = {
    "demo": "模拟盘",
    "live": "实盘"
}

CSV_LOCK = threading.RLock()

# ==================== 动态止盈表（完全复用原OKX表） ====================
PROFIT_STOP_MAP = [
    (0.065, 0.015),
    (0.066, 0.0149),
    (0.067, 0.0148),
    (0.068, 0.0147),
    (0.069, 0.0146),
    (0.070, 0.0145),
    (0.071, 0.0144),
    (0.072, 0.0143),
    (0.073, 0.0142),
    (0.074, 0.0141),
    (0.075, 0.0140),
    (0.076, 0.0139),
    (0.077, 0.0138),
    (0.078, 0.0137),
    (0.079, 0.0136),
    (0.080, 0.0135),
    (0.081, 0.0134),
    (0.082, 0.0133),
    (0.083, 0.0132),
    (0.084, 0.0131),
    (0.085, 0.0130),
    (0.086, 0.0129),
    (0.087, 0.0128),
    (0.088, 0.0127),
    (0.089, 0.0126),
    (0.090, 0.0125),
    (0.091, 0.0124),
    (0.092, 0.0123),
    (0.093, 0.0122),
    (0.094, 0.0121),
    (0.095, 0.0120),
    (0.096, 0.0119),
    (0.097, 0.0118),
    (0.098, 0.0117),
    (0.099, 0.0116),
    (0.100, 0.0115),
    (0.101, 0.0114),
    (0.102, 0.0113),
    (0.103, 0.0112),
    (0.104, 0.0111),
    (0.105, 0.0110),
    (0.106, 0.0109),
    (0.107, 0.0108),
    (0.108, 0.0107),
    (0.109, 0.0106),
    (0.110, 0.0105),
    (0.111, 0.0104),
    (0.112, 0.0103),
    (0.113, 0.0102),
    (0.114, 0.0101),
    (0.115, 0.0100),
    (0.116, 0.0099),
    (0.117, 0.0098),
    (0.118, 0.0097),
    (0.119, 0.0096),
    (0.120, 0.0095),
    (0.121, 0.0094),
    (0.122, 0.0093),
    (0.123, 0.0092),
    (0.124, 0.0091),
    (0.125, 0.0090),
    (0.126, 0.0089),
    (0.127, 0.0088),
    (0.128, 0.0087),
    (0.129, 0.0086),
    (0.130, 0.0085),
    (0.131, 0.0084),
    (0.132, 0.0083),
    (0.133, 0.0082),
    (0.134, 0.0081),
    (0.135, 0.0080),
    (0.136, 0.0079),
    (0.137, 0.0078),
    (0.138, 0.0077),
    (0.139, 0.0076),
    (0.140, 0.0075),
    (0.141, 0.0074),
    (0.142, 0.0073),
    (0.143, 0.0072),
    (0.144, 0.0071),
    (0.145, 0.0070),
    (0.146, 0.0069),
    (0.147, 0.0068),
    (0.148, 0.0067),
    (0.149, 0.0066),
    (0.150, 0.0065),
    (0.151, 0.0064),
    (0.152, 0.0063),
    (0.153, 0.0062),
    (0.154, 0.0061),
    (0.155, 0.0060),
    (0.156, 0.0059),
    (0.157, 0.0058),
    (0.158, 0.0057),
    (0.159, 0.0056),
    (0.160, 0.0055),
    (0.161, 0.0054),
    (0.162, 0.0053),
    (0.163, 0.0052),
    (0.164, 0.0051),
    (0.165, 0.0050),
    (0.166, 0.0049),
    (0.167, 0.0048),
    (0.168, 0.0047),
    (0.169, 0.0046),
    (0.170, 0.0045),
    (0.171, 0.0044),
    (0.172, 0.0043),
    (0.173, 0.0042),
    (0.174, 0.0041),
    (0.175, 0.0040),
    (0.176, 0.0039),
    (0.177, 0.0038),
    (0.178, 0.0037),
    (0.179, 0.0036),
    (0.180, 0.0035),
    (0.181, 0.0034),
    (0.182, 0.0033),
    (0.183, 0.0032),
    (0.184, 0.0031),
    (0.185, 0.0030),
    (0.186, 0.0029),
    (0.187, 0.0028),
    (0.188, 0.0027),
    (0.189, 0.0026),
    (0.190, 0.0025),
    (0.191, 0.0024),
    (0.192, 0.0023),
    (0.193, 0.0022),
    (0.194, 0.0021),
    (0.195, 0.0020),
    (0.196, 0.0019),
    (0.197, 0.0018),
    (0.198, 0.0017),
    (0.199, 0.0016),
    (0.200, 0.0015),
    (0.201, 0.0014),
    (0.202, 0.0013),
    (0.203, 0.0012),
    (0.204, 0.0011),
    (0.205, 0.0010),
    (0.206, 0.0009),
    (0.207, 0.0008),
    (0.208, 0.0007),
    (0.209, 0.0006),
    (0.210, 0.0005),
    (0.211, 0.0004),
    (0.212, 0.0003),
    (0.213, 0.0002),
    (0.214, 0.0001),
    (0.215, 0.0),
    (0.216, 0),
]

def get_stop_loss(profit_pct: float) -> float:
    # 表按利润升序排列，须从高利润端倒序匹配，否则永远命中第一条(0.065,0.015)
    for p, s in reversed(PROFIT_STOP_MAP):
        if profit_pct >= p:
            return s
    return 0.015

# ==================== 日志系统，完全复用原逻辑 ====================
class Logger:
    def __init__(self, log_file: str):
        self.log_file = log_file
        self.ding_webhook = None
        self.ding_secret = None
        self._ensure_log_file()

    def _ensure_log_file(self):
        if not os.path.exists(self.log_file):
            try:
                with open(self.log_file, 'w', encoding='utf-8') as f:
                    f.write(f"# {PROGRAM_NAME} {VERSION} 运行日志\n")
                    f.write(f"# 创建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("=" * 60 + "\n")
            except Exception:
                pass

    def set_dingtalk(self, webhook: str, secret: str = None):
        self.ding_webhook = webhook
        self.ding_secret = secret

    def _should_write_log(self, level: str, category: str) -> bool:
        if level == "ERROR":
            return True
        if level == "INFO":
            return category in ["startup", "shutdown", "config_load", "buy_execute", "sell_execute", "profit_trigger", "stop_loss_trigger", "cooldown_start"]
        if level == "WARN":
            return category in ["insufficient_funds", "connection_error"]
        return False

    def _write(self, level: str, msg: str, category: str = "general"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] [{level}] {msg}"
        print(log_line)
        if self._should_write_log(level, category):
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(log_line + "\n")
            except Exception:
                pass
        if level == "ERROR" and self.ding_webhook:
            self._push_dingtalk_error(msg)

    def _push_dingtalk_error(self, msg: str):
        try:
            push_to_dingtalk(self.ding_webhook, self.ding_secret, f"🚨 系统异常\n\n{msg}\n\n时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as e:
            print(f"[日志] 钉钉异常推送失败: {e}")

    def info(self, msg: str, category: str = "general"):
        self._write("INFO", msg, category)

    def warn(self, msg: str, category: str = "general"):
        self._write("WARN", msg, category)

    def error(self, msg: str, category: str = "general"):
        self._write("ERROR", msg, category)

    def debug(self, msg: str, category: str = "debug"):
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [DEBUG] {msg}")

# ==================== 钉钉推送 ====================
def build_dingtalk_url(webhook: str, secret: str = None) -> str:
    if not secret:
        return webhook
    timestamp = str(int(time.time() * 1000))
    sign_string = f"{timestamp}\n{secret}".encode("utf-8")
    sign = base64.b64encode(hmac.new(secret.encode("utf-8"), sign_string, hashlib.sha256).digest()).decode("utf-8")
    parts = urllib.parse.urlsplit(webhook)
    query = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
    query.extend([("timestamp", timestamp), ("sign", sign)])
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, urllib.parse.urlencode(query), parts.fragment))

def push_to_dingtalk(webhook: str, secret: str = None, text: str = None, title: str = "交易提醒", msg_type: str = "text") -> bool:
    if not webhook:
        return False
    try:
        url = build_dingtalk_url(webhook, secret)
        payload = {"msgtype": "text", "text": {"content": text}} if msg_type == "text" else {"msgtype": "markdown", "markdown": {"title": title, "text": text}}
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json", "User-Agent": "Python DingTalk Bot"}, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8")).get("errcode") in (0, "0", None)
    except Exception as e:
        print(f"[钉钉] 推送失败: {e}")
        return False

# ==================== CSV辅助函数 ====================
def _ensure_csv_header(file_path: str, header: List[str]):
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        try:
            with CSV_LOCK, open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                csv.writer(f).writerow(header)
        except Exception as e:
            print(f"[CSV] 初始化失败 {file_path}: {e}")
        return
    try:
        with CSV_LOCK, open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            existing_header = next(reader, [])
            rows = list(reader)
    except Exception as e:
        print(f"[CSV] 读取 header 失败 {file_path}: {e}")
        return
    if existing_header == header:
        return
    print(f"[CSV] 升级表结构 {file_path}: {existing_header} -> {header}")
    old_idx = {name: i for i, name in enumerate(existing_header)}
    migrated = []
    for row in rows:
        new_row = []
        for col in header:
            if col in old_idx and old_idx[col] < len(row):
                new_row.append(row[old_idx[col]])
            else:
                new_row.append("")
        migrated.append(new_row)
    try:
        with CSV_LOCK, open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(migrated)
    except Exception as e:
        print(f"[CSV] 迁移失败 {file_path}: {e}")

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
        print(f"[CSV] 写入失败 {file_path}: {e}")

def _write_csv_all(file_path: str, header: List[str], rows: List[Dict]):
    try:
        with CSV_LOCK, open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=header)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
    except Exception as e:
        print(f"[CSV] 写入失败 {file_path}: {e}")

# ==================== 黑名单、观察列表 ====================
def load_blacklist() -> set:
    rows = _read_csv(BLACKLIST_FILE, BLACKLIST_HEADER)
    return {row.get("inst_id", "") for row in rows if row.get("inst_id")}

def add_to_blacklist(inst_id: str, reason: str):
    if not inst_id:
        return
    if inst_id in load_blacklist():
        return
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write_csv_row(BLACKLIST_FILE, BLACKLIST_HEADER, {"inst_id": inst_id, "reason": reason, "add_time": now_str})
    print(f"[黑名单] {inst_id} 已加入黑名单，原因: {reason}")

def load_observe_list() -> set:
    rows = _read_csv(OBSERVE_LIST_FILE, OBSERVE_LIST_HEADER)
    return {row.get("inst_id", "") for row in rows if row.get("inst_id")}

def add_to_observe_list(inst_id: str, reason: str):
    if not inst_id:
        return
    if inst_id in load_observe_list():
        return
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write_csv_row(OBSERVE_LIST_FILE, OBSERVE_LIST_HEADER, {"inst_id": inst_id, "reason": reason, "add_time": now_str})
    print(f"[待观察] {inst_id} 已加入待观察列表，原因: {reason}")

# ==================== 多层次价格计算（完全复用OKX数学逻辑） ====================
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

# ==================== FutuOpenD 封装客户端 ====================
class FutuTradeError(Exception):
    pass

class FutuClient:
    def __init__(self, host, port, market, trade_mode, logger: Logger):
        self.host = host
        self.port = port
        self.market = market
        self.trade_mode = trade_mode
        self.logger = logger
        self.quote_ctx: Optional[ft.OpenQuoteContext] = None
        self.trade_ctx: Optional[ft.OpenTradeContext] = None
        self._connect()

    def _connect(self):
        self.quote_ctx = ft.OpenQuoteContext(host=self.host, port=self.port)
        if self.trade_mode == "demo":
            self.trade_ctx = ft.OpenTradeContext(host=self.host, port=self.port, is_simulate=True)
        else:
            if self.market == "CN":
                raise FutuTradeError("A股不支持实盘")
            self.trade_ctx = ft.OpenTradeContext(host=self.host, port=self.port, is_simulate=False)

    def close(self):
        if self.quote_ctx:
            self.quote_ctx.close()
        if self.trade_ctx:
            self.trade_ctx.close()

    def get_acc_balance(self):
        """获取账户可用资金，区分模拟/实盘"""
        ret, data = self.trade_ctx.get_accinfo()
        if ret != ft.RET_OK:
            raise FutuTradeError(f"获取账户信息失败 {data}")
        # 根据市场取对应可用资金；兼容不同 futu 版本的列名
        col = "avail_balance" if "avail_balance" in data.columns else "avail_withdraw_cash"
        return float(data.loc[0, col])

    def get_positions(self):
        ret, data = self.trade_ctx.get_position_list()
        if ret != ft.RET_OK:
            raise FutuTradeError(f"获取持仓失败 {data}")
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

    def market_buy_amount(self, code: str, amount: float, ref_price: float = 0.0):
        """
        按金额市价买入；A股仅模拟打印，不下单
        返回成交结果 dict；error字段表示失败
        """
        if self.market == "CN" and self.trade_mode == "live":
            return {"error": "A股禁止实盘下单，仅模拟"}
        if self.trade_mode == "demo":
            self.logger.info(f"[模拟买入] {code} 拟投入金额 {amount:.2f}")
            return {"price": 0, "qty": 0, "cost": amount, "ord_id": "SIM", "error": None}
        # 实盘：富途API按数量市价下单，需先用参考价折算数量（无按金额市价下单接口）
        price = ref_price if ref_price > 0 else self._get_last_price(code)
        if price <= 0:
            return {"error": f"无法获取参考价折算数量: {code}"}
        qty = int(amount / price)
        if qty <= 0:
            return {"error": f"折算数量为0: 金额{amount:.2f} / 价格{price:.6g}"}
        ret, ret_data = self.trade_ctx.place_order(
            price=0, qty=qty, code=code, order_type=ft.OrderType.MARKET,
            trd_side=ft.TrdSide.BUY, trd_env=ft.TrdEnv.REAL
        )
        if ret != ft.RET_OK:
            return {"error": f"下单失败:{ret_data}"}
        return {"price": 0, "qty": qty, "cost": amount, "ord_id": self._extract_order_id(ret_data), "error": None}

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

    def market_sell_qty(self, code: str, qty: float):
        if self.market == "CN" and self.trade_mode == "live":
            return {"error": "A股禁止实盘下单，仅模拟"}
        if self.trade_mode == "demo":
            self.logger.info(f"[模拟卖出] {code} 拟卖出数量 {qty:.2f}")
            return {"price":0, "qty": qty, "proceeds":0, "ord_id":"SIM", "error": None}
        ret, ret_data = self.trade_ctx.place_order(
            price=0, qty=int(qty), code=code, order_type=ft.OrderType.MARKET,
            trd_side=ft.TrdSide.SELL, trd_env=ft.TrdEnv.REAL
        )
        if ret != ft.RET_OK:
            return {"error": f"卖出失败:{ret_data}"}
        return {"price":0, "qty": qty, "proceeds":0, "ord_id": self._extract_order_id(ret_data), "error": None}

    def subscribe_tick(self, code_list: List[str]):
        ret, err = self.quote_ctx.subscribe(code_list, [ft.SubType.TICKER])
        if ret != ft.RET_OK:
            raise FutuTradeError(f"订阅tick失败:{err}")

# ==================== 全局缓存 ====================
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
    except Exception:
        pass

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
    except Exception:
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
        print(f"[资金] 更新查看文件失败: {e}")

# ==================== 持久化读写函数（与原OKX完全对齐） ====================
def load_positions() -> Dict[str, Dict]:
    rows = _read_csv(POSITIONS_FILE, POSITIONS_HEADER)
    result = {}
    for row in rows:
        inst_id = row.get("inst_id", "")
        if inst_id:
            position_price = float(row.get("position_price", 0) or 0)
            position_qty = float(row.get("position_qty", 0) or 0)
            position_amount = float(row.get("position_amount", 0) or 0)
            raw_cp = row.get("current_price", "") or ""
            current_price = float(raw_cp) if raw_cp else position_price
            mv_raw = row.get("market_value", "") or ""
            upnl_raw = row.get("unrealized_pnl", "") or ""
            upnl_pct_raw = row.get("unrealized_pnl_pct", "") or ""
            if mv_raw != "" and upnl_raw != "":
                market_value = float(mv_raw)
                unrealized_pnl = float(upnl_raw)
                unrealized_pnl_pct = float(upnl_pct_raw) if upnl_pct_raw else 0.0
            else:
                market_value = current_price * position_qty if current_price > 0 else position_amount
                unrealized_pnl = market_value - position_amount if position_amount > 0 else 0.0
                unrealized_pnl_pct = (unrealized_pnl / position_amount * 100.0) if position_amount > 0 else 0.0
            result[inst_id] = {
                "position_price": position_price,
                "position_qty": position_qty,
                "position_amount": position_amount,
                "buy_time": row.get("buy_time", ""),
                "buy_reason": row.get("buy_reason", ""),
                "pair_type": row.get("pair_type", ""),
                "pair_sell_line": row.get("pair_sell_line", ""),
                "peak_price": float(row.get("peak_price", 0) or 0),
                "stop_price": float(row.get("stop_price", 0) or 0),
                "profit_triggered": (str(row.get("profit_triggered") or "False")).strip().lower() == "true",
                "check_level": row.get("check_level", "T3"),
                "last_check_time": row.get("last_check_time", ""),
                "current_price": current_price,
                "market_value": market_value,
                "unrealized_pnl": unrealized_pnl,
                "unrealized_pnl_pct": unrealized_pnl_pct,
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
    _write_csv_row(COOLDOWN_FILE, COOLDOWN_HEADER, {
        "inst_id": inst_id,
        "sell_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cooldown_until": cooldown_until
    })

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

# ==================== 配置加载 ====================
def load_config() -> Tuple[Dict, Dict]:
    import openpyxl
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(f"配置文件不存在: {CONFIG_FILE}")
    config = {}
    # 资金口径：一律以富途API账户可用余额为准，配置中禁止设置资金字段（total_funds 已移除，防止与真实账户不一致）
    system_params = {
        "buy_ratio": 0.01,
        "cooldown_days": 7,
        "profit_trigger": 0.05,
        "stop_loss": 0.015,
        "profit_step": 0.01,
        "break_alert_max": 4,
        "break_alert_interval": 1020,
        "push_interval": 3.0,
        "data_source": "FUTU",
        "trade_mode": "demo"
    }
    wb = openpyxl.load_workbook(CONFIG_FILE, data_only=True)
    if "币对配置" in wb.sheetnames:
        ws = wb["币对配置"]
        for row in range(2, ws.max_row + 1):
            inst_id = ws.cell(row=row, column=1).value
            ref_high = ws.cell(row=row, column=2).value
            ref_low = ws.cell(row=row, column=3).value
            enabled = ws.cell(row=row, column=4).value
            if inst_id and enabled and str(enabled).upper() in ("YES", "是", "1"):
                try:
                    ref_high = float(ref_high)
                    ref_low = float(ref_low)
                except (TypeError, ValueError):
                    continue
                config[inst_id] = {"ref_high": ref_high, "ref_low": ref_low}
    if "系统参数" in wb.sheetnames:
        ws = wb["系统参数"]
        for row in range(2, ws.max_row + 1):
            key = ws.cell(row=row, column=1).value
            value = ws.cell(row=row, column=2).value
            if key and value is not None:
                key = str(key).strip()
                if key in system_params:
                    if isinstance(system_params[key], float):
                        system_params[key] = float(value)
                    elif isinstance(system_params[key], int):
                        system_params[key] = int(float(value))
                    else:
                        system_params[key] = str(value)
    env_cooldown = os.getenv("COOLDOWN_DAYS")
    if env_cooldown is not None:
        try:
            system_params["cooldown_days"] = int(env_cooldown)
        except Exception:
            pass
    return config, system_params

# ==================== SymbolState：业务逻辑完全复刻OKX版本 ====================
class SymbolState:
    def __init__(self, inst_id, ref_high, ref_low, logger, ding_webhook=None, ding_secret=None, system_params=None):
        self.inst_id = inst_id
        self.logger = logger
        self.ding_webhook = ding_webhook
        self.ding_secret = ding_secret
        self.params = system_params or {}
        # 资金一律取富途API账户余额(get_futu_balance)，不设配置文件资金字段
        self.min_buy_amount = float(os.getenv("MIN_BUY_AMOUNT", "100"))
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
        self._okx_available = True

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
                    self.logger.info(f"{inst_id} API检测持仓，本地同步更新", "startup")
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
            self.logger.info(f"{self.inst_id} 账户余额 {balance:.2f} 低于最小买入金额 {self.min_buy_amount:.2f}，跳过买入并冷却300秒")
            return None
        buy_level = getattr(self.levels, buy_key)
        if price >= buy_level:
            self.logger.debug(f"{self.inst_id} 当前价 {price:.8g} 不小于买点 {buy_level:.8g}，跳过")
            return None
        can, msg = self._can_buy()
        if not can:
            self.logger.debug(f"{self.inst_id} 买入跳过: {msg}")
            return None
        raw_amount = balance * self.params.get("buy_ratio", 0.01)
        buy_amount = max(raw_amount, self.min_buy_amount)
        if buy_amount > balance:
            buy_amount = balance
        fill = self.client.market_buy_amount(self.inst_id, buy_amount, ref_price=price)
        now = time.time()
        if fill is None:
            self.logger.error(f"{self.inst_id} API买入失败返回None")
            return self._handle_unknown_error(reason="未知错误返回None")
        if isinstance(fill, dict) and fill.get("error"):
            error_msg = fill["error"]
            self.logger.error(f"{self.inst_id} API买入失败:{error_msg}")
            return self._handle_unknown_error(reason=error_msg[:100])
        # 成交成功
        price = fill["price"] if fill["price"] != 0 else price
        qty = fill["qty"] if fill["qty"] != 0 else (buy_amount / price if price > 0 else 0)
        if qty <= 0:
            self.logger.error(f"{self.inst_id} 买入成交数量异常 qty={qty}，忽略该笔")
            return None
        buy_amount = fill["cost"]
        reason = f"{reason} | ordId={fill.get('ord_id','')}"
        self.logger.info(f"{self.inst_id} 买入成交 ordId={fill.get('ord_id')} 均价 {price:.8g} 数量 {qty:.8g} 花费 {buy_amount:.2f}", "buy_execute")
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
        self.logger.info(msg, "buy_execute")
        self._push_dingtalk(f"💰 买入\n\n{self.inst_id}\n{pair_name}买点: {buy_key}\n价格: {price:.8g}\n金额: {buy_amount:.2f}")
        return msg

    def _handle_unknown_error(self, reason: str):
        now = time.time()
        if now - self._first_fail_time > self._window_duration:
            self._first_fail_time = now
            self._fail_count_in_window = 1
            self._retry_after = now + self._retry_delay
            self.logger.info(f"{self.inst_id} 未知错误第1次，{self._retry_delay}秒后重试")
            return None
        else:
            self._fail_count_in_window += 1
            if self._fail_count_in_window >= 2:
                self.logger.warn(f"{self.inst_id} 连续两次未知错误，加入待观察列表")
                add_to_observe_list(self.inst_id, f"未知错误连续失败:{reason[:100]}")
                self._remove_from_buy_queue()
                self._first_fail_time = 0.0
                self._fail_count_in_window = 0
                self._retry_after = 0.0
                return None
            else:
                self._retry_after = now + self._retry_delay
                self.logger.info(f"{self.inst_id} 未知错误第{self._fail_count_in_window}次，{self._retry_delay}秒后重试")
                return None

    def _remove_from_buy_queue(self):
        self.buy_queue = [r for r in self.buy_queue if r.get("inst_id") != self.inst_id]
        save_buy_queue(self.buy_queue)

    def _execute_sell(self, price, reason):
        can, msg = self._can_sell()
        if not can:
            self.logger.debug(f"{self.inst_id} 卖出跳过: {msg}")
            return None
        qty = self.position_qty
        fill = self.client.market_sell_qty(self.inst_id, qty)
        if not fill or fill.get("error"):
            self.logger.error(f"{self.inst_id} API卖出失败 {fill}")
            self._push_dingtalk(f"⚠️卖出失败\n{self.inst_id}\n{reason}\n数量:{qty:.8g}")
            return None
        price = fill["price"] if fill["price"] != 0 else price
        qty = fill["qty"]
        amount = fill["proceeds"] if fill["proceeds"] != 0 else price * qty
        reason = f"{reason} | ordId={fill.get('ord_id','')}"
        self.logger.info(f"{self.inst_id} 卖出成交 ordId={fill.get('ord_id')} 均价 {price:.8g} 数量 {qty:.8g} 到手 {amount:.2f}", "sell_execute")
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
        self.logger.info(msg, "sell_execute")
        self._push_dingtalk(f"💸 卖出\n\n{self.inst_id}\n{reason}\n价格: {price:.8g}\n盈亏: {profit:.2f} ({profit_pct:.2f}%)")
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
            self.logger.info(f"{self.inst_id} 盈利≥20%立即清仓", "profit_trigger")
            return self._execute_sell(price, f"盈利率 {profit_pct*100:.1f}% ≥20%，清仓")
        new_stop = self.peak_price * (1 - stop_ratio)
        if new_stop > self.stop_price:
            self.stop_price = new_stop
        if profit_pct >= 0.05 and not self.profit_triggered:
            self.profit_triggered = True
            self.logger.info(f"{self.inst_id} 盈利率 {profit_pct*100:.2f}% ≥5%，启动动态止盈，当前止损价 {self.stop_price:.8g}", "profit_trigger")
        if self.profit_triggered and price <= self.stop_price:
            self.logger.info(f"{self.inst_id} 触发止盈清仓: 当前价 {price:.8g} ≤ 止盈价 {self.stop_price:.8g}", "stop_loss_trigger")
            return self._execute_sell(price, f"止盈清仓 (峰值 {self.peak_price:.8g}, 止损 {stop_ratio*100:.1f}%)")
        return None

    def update_price(self, price, ts):
        alerts = []
        if price == self._last_processed_price and ts == self._last_processed_ts:
            return alerts
        self._last_processed_price = price
        self._last_processed_ts = ts
        self._sync_buy_queue()

        # 更新高低点
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

        # 持仓处理：止盈、配对卖点卖出
        if self.has_position:
            profit_pct = (price - self.position_price) / self.position_price if self.position_price > 0 else 0
            self._update_check_level(profit_pct)
            if self.check_level == "T0":
                res = self._check_take_profit(price)
                if res:
                    alerts.append(f"止盈清仓 @ {price:.8g}")
                    self._save_state()
                    return alerts
            # 配对卖点触发
            if self.is_paired and not self.is_sold:
                sell_price = self.pair_sell_line_price
                if price >= sell_price:
                    res = self._execute_sell(price, f"{self.pair_type}卖点 {self.pair_sell_line} 触发")
                    if res:
                        alerts.append(f"{self.pair_type}卖点清仓 @ {price:.8g}")
                        self._save_state()
                        return alerts

        # 无持仓：反弹买入信号逻辑
        if not self.has_position and not self._is_in_cooldown() and self.inst_id not in self.observe_list and self.inst_id not in self.blacklist:
            buy_key, buy_level = self._get_buy_level(price)
            if buy_key is not None:
                if not self.buy_signal_activated:
                    self.buy_signal_activated = True
                    self.buy_signal_lowest_price = price
                    self.buy_signal_buy_level = buy_key
                    self.buy_signal_buy_price = buy_level
                    self.buy_signal_attempt_time = 0.0
                    self.logger.debug(f"{self.inst_id} 买入信号激活，买点 {buy_key} ({buy_level:.8g})，当前价 {price:.8g}")
                else:
                    if price < self.buy_signal_lowest_price:
                        self.buy_signal_lowest_price = price
                        self.logger.debug(f"{self.inst_id} 更新最低点 {price:.8g}")
            else:
                if self.buy_signal_activated:
                    if price >= self.buy_signal_buy_price:
                        self.logger.debug(f"{self.inst_id} 价格回升至买点之上，取消买入信号")
                        self.buy_signal_activated = False
                        self.buy_signal_lowest_price = 0.0
                        self.buy_signal_buy_level = ""
                        self.buy_signal_buy_price = 0.0

            if self.buy_signal_activated:
                now = time.time()
                if now - self.buy_signal_attempt_time >= 30:
                    rebound_price = self.buy_signal_lowest_price * 1.01
                    if price >= rebound_price and price < self.buy_signal_buy_price:
                        self.logger.debug(f"{self.inst_id} 满足反弹买入条件：最低 {self.buy_signal_lowest_price:.8g}, 反弹1%={rebound_price:.8g}, 当前价 {price:.8g}")
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

    def _push_dingtalk(self, text):
        if self.ding_webhook:
            push_to_dingtalk(self.ding_webhook, self.ding_secret, text)


# ==================== 行情引擎，Futu TICK 回调 ====================
class MarketEngine:
    def __init__(self, logger, ding_webhook=None, ding_secret=None, system_params=None):
        self.logger = logger
        self.ding_webhook = ding_webhook
        self.ding_secret = ding_secret
        self.params = system_params or {}
        self.states: Dict[str, SymbolState] = {}
        self.lock = threading.Lock()

    def add_state(self, inst_id, ref_high, ref_low):
        state = SymbolState(inst_id, ref_high, ref_low, self.logger, self.ding_webhook, self.ding_secret, self.params)
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
                self.logger.info(f"[{inst_id}] " + " | ".join(key_alert), "trade_event")
            else:
                self.logger.debug(f"[{inst_id}] " + " | ".join(alerts))


# tick 回调函数
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


market_engine: Optional[MarketEngine] = None

# ==================== 主函数 ====================
def main():
    global _futu_client, market_engine
    print("=" * 60)
    print(f"{PROGRAM_NAME} {VERSION}")
    print(f"工作目录: {BASE_DIR}")
    print(f"数据目录: {DATA_DIR}")
    print(f"运行市场: {RUN_MARKET} | 模式: {MODE_DESC[TRADE_MODE]}")
    print("=" * 60)

    logger = Logger(LOG_FILE)
    ding_webhook = os.getenv("DINGTALK_WEBHOOK_futu")
    ding_secret = os.getenv("DINGTALK_SECRET_futu")
    if ding_webhook:
        logger.set_dingtalk(ding_webhook, ding_secret)
        logger.info("钉钉推送已启用", "startup")
    else:
        logger.warn("钉钉推送未配置，仅终端输出", "startup")

    # 加载Excel配置
    try:
        config, system_params = load_config()
        logger.info(f"加载配置成功: {len(config)} 个标的", "config_load")
    except Exception as e:
        logger.error(f"配置加载失败: {e}", "config_load")
        sys.exit(1)
    if not config:
        logger.error("Excel配置中没有启用的标的", "config_load")
        sys.exit(1)

    # 初始化富途客户端
    try:
        _futu_client = FutuClient(FUTU_OPEND_HOST, FUTU_OPEND_PORT, RUN_MARKET, TRADE_MODE, logger)
        bal = _futu_client.get_acc_balance()
        logger.info(f"FutuOpenD 连接成功，账户可用资金: {bal:.2f}", "startup")
        _update_funds_for_view(bal)
        refresh_positions_cache(force=True)
    except Exception as e:
        logger.error(f"连接FutuOpenD失败: {e}", "startup")
        print("\n提示：请确认FutuOpenD已启动、账号已登录；云服务器需要内网穿透指向本地FutuOpenD")
        sys.exit(1)

    min_amt = float(os.getenv("MIN_BUY_AMOUNT", "100"))
    cooldown_days = int(os.getenv("COOLDOWN_DAYS", system_params.get("cooldown_days", 7)))
    print(f"单笔最小买入金额: {min_amt:.2f}")
    print(f"交易静默期天数: {cooldown_days}")

    market_engine = MarketEngine(logger, ding_webhook, ding_secret, system_params)
    blacklist = load_blacklist()
    observe_list = load_observe_list()
    if blacklist:
        logger.info(f"黑名单数量: {len(blacklist)}", "startup")
    if observe_list:
        logger.info(f"待观察列表数量: {len(observe_list)}", "startup")

    code_list = list(config.keys())
    logger.info("初始化标的状态...", "startup")
    for inst_id, data in config.items():
        market_engine.add_state(inst_id, data["ref_high"], data["ref_low"])
        st = market_engine.states[inst_id]
        print(f"\n{inst_id}")
        print(f"  最高:{st.high:.6g} 最低:{st.low:.6g}")
        print(f"  回撤上:{st.levels.pullback_up:.6g} 回撤2下:{st.levels.pullback2_down:.6g}")
        print(f"  回调上:{st.levels.callback_up:.6g} 回调2下:{st.levels.callback2_down:.6g}")
        print(f"  过渡上:{st.levels.transition_up:.6g} 过渡2下:{st.levels.transition2_down:.6g}")
        print(f"  极限上:{st.levels.limit_up:.6g} 极限2下:{st.levels.limit2_down:.6g}")

    # 订阅行情tick
    try:
        _futu_client.subscribe_tick(code_list)
        _futu_client.quote_ctx.set_handler(tick_callback)
        logger.info(f"行情订阅完成，共{len(code_list)}个标的", "startup")
    except Exception as e:
        logger.error(f"订阅tick行情失败:{e}", "startup")
        sys.exit(1)

    api_pos = get_api_positions()
    if api_pos:
        print("\n==== 当前富途账户持仓 ====")
        for code, pos_data in api_pos.items():
            print(f"  {code}: 数量 {pos_data['position_qty']:.6g}, 成本价 {pos_data['position_price']:.6g}")

    local_pos = load_positions()
    for inst, d in local_pos.items():
        print(f"本地恢复持仓: {inst} qty={d['position_qty']:.6g} @ {d['position_price']:.6g}")

    buy_queue = load_buy_queue()
    pending = [x for x in buy_queue if x.get("status") == "待买入"]
    print(f"\n待买入队列数量: {len(pending)}")
    print(f"账户可用资金: {get_futu_balance(force_refresh=True):.2f}")
    logger.info("="*60, "startup")
    logger.info("策略运行中，Ctrl+C停止程序", "startup")
    logger.info("买入规则：跌破买点激活信号，从最低点反弹1%且仍低于买点执行买入", "startup")
    logger.info("="*60, "startup")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("收到Ctrl+C，程序准备退出", "shutdown")
    finally:
        if _futu_client:
            _futu_client.close()
        _update_funds_for_view(get_futu_balance(force_refresh=True))
        logger.info("程序已退出", "shutdown")


if __name__ == "__main__":
    main()
