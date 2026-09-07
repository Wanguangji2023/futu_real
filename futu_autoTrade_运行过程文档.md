# futu_autoTrade 运行过程文档（v1.1）

> 适用文件：`futu_real.py`（v1.3，2026-09-07 修复版）
> 程序定位：基于 FutuOpenD 的 A股/港股/美股自动化交易策略，复刻原 OKX 版「跌破买点 → 反弹1%买入 → 配对卖点/动态移动止盈卖出」逻辑。
> 本文档逐层拆解程序从启动到退出的完整运行过程，含配置、状态机、每步判定条件、持久化与异常处理。

---

## 目录

1. [程序概述与架构](#1-程序概述与架构)
2. [运行环境与前置条件](#2-运行环境与前置条件)
3. [配置体系](#3-配置体系)
4. [启动流程逐步分解](#4-启动流程逐步分解)
5. [核心数据模型](#5-核心数据模型)
6. [行情驱动主流程（update_price 全流程）](#6-行情驱动主流程update_price-全流程)
7. [买入逻辑详解](#7-买入逻辑详解)
8. [卖出逻辑详解](#8-卖出逻辑详解)
9. [冷却、黑名单与观察列表](#9-冷却黑名单与观察列表)
10. [持久化体系（CSV 全字段）](#10-持久化体系csv-全字段)
11. [异常处理与告警](#11-异常处理与告警)
12. [已知限制与注意点](#12-已知限制与注意点)
13. [v1.1 修复记录](#13-v11-修复记录)

---

## 1. 程序概述与架构

### 1.1 定位

程序通过 **FutuOpenD 网关** 获取实时 TICK 行情并下单，对配置表中的每个标的维护独立状态机：

- **买入侧**：价格跌破任意「2下」买点 → 激活买入信号并记录最低点 → 价格自最低点**反弹 1%** 且仍低于买点 → 执行买入。
- **卖出侧**（三选一先到先出）：
  1. 配对卖点（固定止盈位，如「回撤上沿」）；
  2. 动态移动止盈（盈利 ≥5% 启动，回撤比例随利润升高从 1.5% 递减到 0）；
  3. **盈利 ≥20% 无条件立即清仓**（v1.1 调整）。
- 卖出后进入静默期（默认 7 天），到期自动恢复可买。
- A股强制模拟盘；港股/美股支持模拟与实盘（`.env` 开关控制）。

### 1.2 分层架构

```
┌─────────────────────────────────────────────────────────────┐
│  FutuOpenD 网关（127.0.0.1:11111，需本地已启动并登录）        │
└───────────────┬──────────────────────────┬───────────────────┘
                │ TICK 行情推送（独立线程）  │ 交易/账户查询（同步调用）
┌───────────────▼──────────────────────────▼───────────────────┐
│  FutuClient 封装层                                             │
│   - OpenQuoteContext 行情（订阅TICKER、快照取价）              │
│   - OpenTradeContext 交易（账户资金/持仓/下单，demo/live）     │
└───────────────┬───────────────────────────────────────────────┘
┌───────────────▼───────────────────────────────────────────────┐
│  MarketEngine（行情分发）                                      │
│   lock 保护 → 按 inst_id 路由到对应 SymbolState                 │
└───────────────┬───────────────────────────────────────────────┘
┌───────────────▼───────────────────────────────────────────────┐
│  SymbolState（每个标的独立状态机，核心业务逻辑）                │
│   PriceLevels（层次价格）· 买入信号 · 止盈/卖点 · 冷却          │
└───────────────┬───────────────────────────────────────────────┘
┌───────────────▼───────────────────────────────────────────────┐
│  持久化层（CSV_LOCK 保护）· 钉钉推送 · 日志                    │
│  tradeRecord / positions / cooldown / high_low / pair_status  │
│  buy_queue / funds / run.log · blacklist / observe_list       │
└───────────────────────────────────────────────────────────────┘
```

### 1.3 线程模型

| 线程 | 职责 | 同步机制 |
|---|---|---|
| 主线程 | 启动初始化 → `while True: sleep(1)` 保活 | 阻塞等待，Ctrl+C 退出 |
| 行情推送线程（Futu 内部） | `tick_callback` → `market_engine.on_tick` | `MarketEngine.lock`（每个 tick 加锁） |
| 文件 I/O | 所有 CSV 读写 | `CSV_LOCK`（`threading.RLock`，可重入） |

> 注意：`tick_callback` 里的 `market_engine` 是模块级全局变量，在 `main()` 中赋值，推送线程读取全局可见。

---

## 2. 运行环境与前置条件

### 2.1 前置条件

| 条件 | 说明 |
|---|---|
| FutuOpenD | 已安装并启动，端口默认 11111（可在 .env 改） |
| 富途账户登录 | FutuOpenD 客户端需已登录；**demo 模式需登录模拟账户**，live 需登录真实账户 |
| 云服务器场景 | 若程序跑在云服务器，需用内网穿透把 FutuOpenD 端口映射到本机 |
| Python 依赖 | `futu`（富途官方 SDK）、`openpyxl`（读配置）、`python-dotenv`（读 .env） |

### 2.2 目录结构

```
D:\A-newWarehouse\FUTU\futu_real\
├── futu_real.py                  # 程序本体（v1.1）
├── .env                          # 环境变量（可选，不存在则用默认值）
├── futu_autoTrade_config.xlsx    # 标的与系统参数配置（必需）
├── futu_autoTrade_blacklist.csv      # 黑名单（手工维护）
├── futu_autoTrade_observe_list.csv   # 观察列表（程序可自动追加）
└── futu_autoTrade_data\          # 运行数据目录（自动创建）
    ├── futu_autoTrade_tradeRecord.csv  # 成交记录
    ├── futu_autoTrade_positions.csv    # 本地持仓状态
    ├── futu_autoTrade_cooldown.csv     # 静默期
    ├── futu_autoTrade_high_low.csv     # 历史最高/最低
    ├── futu_autoTrade_pair_status.csv  # 配对买卖状态
    ├── futu_autoTrade_buy_queue.csv    # 买入队列
    ├── futu_autoTrade_funds.csv        # 资金快照（查看用）
    └── futu_autoTrade_run.log          # 运行日志
```

### 2.3 启动命令

```bash
cd D:\A-newWarehouse\FUTU\futu_real
python futu_real.py
```

> 请勿在本机系统 Python 直接运行（缺 dotenv/futu），需在配置好依赖的虚拟环境中执行。

---

## 3. 配置体系

### 3.1 环境变量（.env）

| 变量名 | 默认值 | 说明 |
|---|---|---|
| `FUTU_OPEND_HOST` | `127.0.0.1` | FutuOpenD 地址 |
| `FUTU_OPEND_PORT` | `11111` | FutuOpenD 端口 |
| `MARKET` | `HK` | 运行市场：`CN`=A股 / `HK`=港股 / `US`=美股 |
| `TRADE_MODE` | `demo` | `demo`=模拟盘 / `live`=实盘（**CN 强制 demo，忽略此值**） |
| `MIN_BUY_AMOUNT` | `100` | 单笔最小买入金额 |
| `COOLDOWN_DAYS` | Excel 值，缺省 `7` | 卖出后静默天数 |
| `DINGTALK_WEBHOOK_futu` | 无 | 钉钉机器人 Webhook（推送买卖/错误） |
| `DINGTALK_SECRET_futu` | 无 | 钉钉加签密钥（可选） |

> 容错说明（v1.3）：数值型环境变量（端口、金额、天数）读取时自动清理复制粘贴残留字符（尾部 `|`、空格、引号），清理后仍无法解析则回退默认值并打印告警，不再导致启动崩溃。例：`FUTU_OPEND_PORT=11111|` 会被正确解析为 `11111`。但**建议仍修正 .env 源文件**，避免隐患。

### 3.2 Excel 配置文件（futu_autoTrade_config.xlsx）

**Sheet 1：`币对配置`**（每行一个标的）

| 列 | 含义 | 取值 |
|---|---|---|
| A | inst_id | 富途代码，如 `HK.00700`、`US.AAPL`、`SH.600519` |
| B | ref_high | 初始参考最高价（首次运行时使用；此后以程序记录的高低价为准） |
| C | ref_low | 初始参考最低价 |
| D | enabled | 启用标记：`YES` / `是` / `1` 生效；空或其他值跳过 |

> 空值容错：B/C 列空或非数字时该行自动跳过（v1.1 修复），不再崩溃。

**Sheet 2：`系统参数`**（key / value 两列）

| 参数 | 默认值 | 实际生效情况 |
|---|---|---|
| `total_funds` | 1000000 | ✅ **已从配置项中移除（v1.2）**：资金一律以富途平台 API 的账户可用余额为准；Excel 中即使写了该参数也不会被读取，从源头杜绝配置与真实账户不一致 |
| `buy_ratio` | 0.01 | ✅ 单笔买入 = 账户余额 × 比例 |
| `cooldown_days` | 7 | ✅ 卖出后静默天数（.env 的 COOLDOWN_DAYS 可覆盖） |
| `profit_trigger` | 0.05 | ❌ 硬编码 0.05（T0 门槛），参数不生效 |
| `stop_loss` | 0.015 | ❌ 硬编码（get_stop_loss 兜底值），参数不生效 |
| `profit_step` / `break_alert_max` / `break_alert_interval` / `push_interval` / `data_source` / `trade_mode` | — | ❌ 保留字段，当前逻辑未使用 |

### 3.3 市场 × 模式矩阵

| 市场 | demo | live |
|---|---|---|
| CN（A股） | ✅ 支持 | ❌ 强制转 demo（`_connect` 直接抛错拒绝实盘） |
| HK | ✅ | ✅（`OpenTradeContext(is_simulate=False)`） |
| US | ✅ | ✅ |

---

## 4. 启动流程逐步分解

`main()` 的完整执行序列（任一步失败 → 打印错误 → `sys.exit(1)`）：

```
[1] 打印程序横幅（版本/工作目录/数据目录/市场/模式）
    ↓
[2] 初始化 Logger（创建 run.log），配置钉钉（有 webhook 才启用）
    ↓
[3] load_config() 读取 Excel
    ├─ 文件不存在 → 报错退出
    ├─ 配置解析失败 → 报错退出
    └─ 无启用标的 → 报错退出
    ↓
[4] FutuClient 连接 FutuOpenD（_connect）
    ├─ quote_ctx = OpenQuoteContext(host, port)
    └─ trade_ctx = OpenTradeContext(host, port, is_simulate)
       （CN 且 live → 直接抛 FutuTradeError）
    ↓
[5] 取账户可用资金 → 写 funds.csv 快照 → 刷新持仓缓存
    连接/取资失败 → 提示检查 FutuOpenD 与登录 → 退出
    ↓
[6] 读取 MIN_BUY_AMOUNT、COOLDOWN_DAYS 并打印
    ↓
[7] 创建 MarketEngine；加载黑名单、观察列表并打印数量
    ↓
[8] 遍历配置逐个 add_state() → 每个标的初始化 SymbolState
    （详见 4.1）并打印各标的层次价格
    ↓
[9] subscribe_tick(code_list) 订阅 TICKER + set_handler(tick_callback)
    订阅失败 → 退出
    ↓
[10] 打印 API 真实持仓 / 本地恢复持仓 / 待买入队列数量 / 可用资金
    ↓
[11] 进入主循环 while True: time.sleep(1)（保活，逻辑全由 tick 驱动）
    ↓
[Ctrl+C] → 关闭 quote/trade 上下文 → 更新 funds 快照 → 记录退出日志
```

### 4.1 SymbolState 初始化（add_state 内部细节）

每个标的创建状态时依次执行：

1. 加载五张状态表：`positions / cooldown / high_low / pair_status / buy_queue`；
2. 恢复高低点：优先用 `high_low` 表中持久化的值，无记录则用 Excel 的 `ref_high/ref_low`；
3. 用恢复的高低点构造 `PriceLevels`，计算全部层次价格；
4. 恢复持仓字段（价格/数量/金额/买点/峰值/止损线/止盈标志/检查级别等）；
5. 恢复配对状态（是否配对、是否已卖、卖点价、盈亏）；
6. 恢复冷却期、买入队列；
7. **API 持仓对账**：若 Futu 账户实际持有该标的而本地无记录 → 以 API 数据补建本地持仓并落盘（防丢单场景）；
8. `_sync_buy_queue()` 同步队列状态。

---

## 5. 核心数据模型

### 5.1 PriceLevels 层次价格（完全复刻 OKX 数学逻辑）

以 `high`（历史最高）与 `low`（历史最低）为锚点，逐级二分构造四个区域、各 4 个关键位：

| 区域 | 关键位 | 公式 |
|---|---|---|
| 公共 | high_mid / mid / low_mid | `(high+low)/2` 的上下半区中点 |
| 回撤区（上方） | pullback1_mid / pullback_up / pullback2_down / pullback3 / pullback4 | 以 `high_mid`~`high` 区间逐级二分；`pullback_up`=(high+pullback1_mid)/2，`pullback2_down`=(pullback1_mid+high_mid)/2 |
| 回调区（中上） | callback1_mid / callback_up / callback2_down / callback3 / callback4 | 以 `high_mid`~`mid` 区间逐级二分 |
| 过渡区（中下） | transition1_mid / transition_up / transition2_down / transition3 / transition4 | 以 `mid`~`low_mid` 区间逐级二分 |
| 极限区（下方） | limit1_mid / limit_up / limit2_down / limit3 / limit4 | 以 `low_mid`~`low` 区间逐级二分 |

**买卖点配对**（每对 = 同区域的「2下」买点 + 「上」卖点）：

| 买点（跌破激活） | 对应卖点（升破卖出） | 名称 |
|---|---|---|
| pullback2_down | pullback_up | 回撤 |
| callback2_down | callback_up | 回调 |
| transition2_down | transition_up | 过渡 |
| limit2_down | limit_up | 极限 |

> 卖点恒高于对应买点（几何保证），因此配对卖点卖出必然是盈利退出。
> 高低点一旦刷新（创新高/新低），所有层次价格随 `levels.update(high, low)` 全量重算。

### 5.2 SymbolState 关键状态字段

| 分组 | 字段 |
|---|---|
| 持仓 | `has_position / position_price / position_qty / position_amount / buy_time / buy_reason` |
| 配对 | `pair_type / pair_sell_line / is_paired / pair_buy_price / pair_sell_line_price / is_sold / sell_time / sell_reason` |
| 止盈 | `peak_price（峰值）/ stop_price（止损线）/ profit_triggered / check_level（T0-T3）/ last_check_time` |
| 买入信号 | `buy_signal_activated / buy_signal_lowest_price / buy_signal_buy_level / buy_signal_buy_price / buy_signal_attempt_time` |
| 冷却/容错 | `cooldown / _balance_insufficient_until / _retry_after / _fail_count_in_window / _first_fail_time` |

### 5.3 全局缓存（减少 API 调用）

| 缓存 | TTL | 说明 |
|---|---|---|
| 账户余额 `_balance_cache` | 3 秒 | 买入判定高频读取时复用 |
| 账户持仓 `_positions_cache` | 30 秒 | `has_api_position` 判重用 |

---

## 6. 行情驱动主流程（update_price 全流程）

每个 tick 到达后 `MarketEngine.on_tick` 加锁并调用 `state.update_price(price, ts)`，完整分支如下：

```
update_price(price, ts)
│
├─[去重] price 与 ts 均与上一条相同 → 直接返回（不重复处理）
│
├─[队列同步] _sync_buy_queue()：状态变化才写盘
│
├─[高低点] 
│   价格 > high → 更新 high、重算层次、写 high_low.csv、记"创新高"
│   价格 < low  → 更新 low、重算层次、写 high_low.csv、记"创新低"
│
├─[有持仓分支]
│   profit_pct = (price - position_price) / position_price
│   _update_check_level(profit_pct)：
│       ≥5%   → T0
│       0~5%  → T1
│       ≤0 且配对类型在(回撤/回调/过渡/极限) → T2
│       否则   → T3
│   │
│   ├─ check_level == T0 → _check_take_profit(price)
│   │      ├─ 触发(≥20%清仓 / 跌破止损线) → 卖出 → 保存 → 返回
│   │      └─ 未触发 → 继续
│   │
│   └─ 配对卖点：price ≥ pair_sell_line_price → 卖出 → 保存 → 返回
│
└─[无持仓分支]（且不在冷却、不在观察列表、不在黑名单）
    buy_key, buy_level = _get_buy_level(price)   # 找 price 之下的最低买点
    │
    ├─ 存在买点（price < 某买点）：
    │     信号未激活 → 激活：记录最低点=price、买点、buy_price，attempt_time=0
    │     信号已激活 → price < 最低点 → 更新最低点
    │
    ├─ 不存在买点：
    │     信号已激活 且 price ≥ buy_price → 取消信号（价格回升买点之上）
    │
    └─ 信号激活状态：
        now - attempt_time ≥ 30 秒 且
        price ≥ 最低点×1.01（反弹1%）且 price < buy_price
        → _execute_buy() 执行买入
        → 成交则保存并返回；未成交则记录 attempt_time 后 30 秒再试
```

**T0/T1/T2/T3 的含义**：T0 才评估动态止盈；T1~T3 只评估配对卖点；T2 专用于"微亏但已配对"的正常持仓态。

---

## 7. 买入逻辑详解

### 7.1 买入信号状态机（三个状态）

```
  ┌─────────┐   price 跌破任意买点   ┌─────────────┐
  │ 未激活   │ ────────────────────▶ │ 已激活       │
  └─────────┘                       └─────────────┘
      ▲                                  │  price 回升至 ≥ buy_price
      │                                  ▼  （取消信号，回到未激活）
      └──────────── price ≥ buy_price ───┘
      已激活状态内：price 创新低 → 最低点下移（追跌更便宜）
```

### 7.2 买点选择规则（_get_buy_level）

按 `pullback2_down → callback2_down → transition2_down → limit2_down` 顺序扫描，返回 **price 之下最低的那个买点**（四个买点逐级降低，因此等价于"当前价跌破了哪一档"）：

| 当前价位置 | 选中的买点 |
|---|---|
| `pullback2_down > 价 ≥ callback2_down` | pullback2_down（回撤） |
| `callback2_down > 价 ≥ transition2_down` | callback2_down（回调） |
| `transition2_down > 价 ≥ limit2_down` | transition2_down（过渡） |
| `价 < limit2_down` | limit2_down（极限） |

### 7.3 反弹买入条件（三条件同时满足）

1. 信号已激活（此前价格跌破过买点）；
2. `now - buy_signal_attempt_time ≥ 30 秒`（节流，防止同一波下跌内反复尝试）；
3. `当前价 ≥ 最低点 × 1.01`（自最低点反弹 ≥1%）**且** `当前价 < 买点价`（仍在买点之下）。

> 含义：不接下跌中的飞刀，等跌势企稳反弹 1% 确认后再入场，同时要求买入价仍比初始买点便宜。

### 7.4 _can_buy 全部拦截条件（任一命中则放弃本次买入）

| # | 条件 | 提示文案 |
|---|---|---|
| 1 | 本地已有持仓 | 本地已有持仓 |
| 2 | API 检测到现有持仓 | API检测到现有持仓 |
| 3 | 处于静默期 | 静默期中 (至 ...) |
| 4 | 在观察列表 | 在待观察列表中 |
| 5 | 在黑名单 | 在黑名单中（v1.1 修复） |
| 6 | 余额不足冷却中（300 秒） | 余额不足冷却中 |
| 7 | 未知错误重试延迟中 | 未知错误重试延迟中（剩余 N 秒） |
| 8 | 余额 < 最小买入金额 | 资金不足（并启动 300 秒冷却） |
| 9 | 已配对未卖出 | 已配对未卖出 |

### 7.5 _execute_buy 执行细节

```
1. 强制刷新余额；余额 < MIN_BUY_AMOUNT → 跳过并启动 300 秒余额冷却
2. 防御：当前价 ≥ 买点 → 跳过（与信号层重复校验）
3. _can_buy() 全部检查
4. 金额 = max(余额 × buy_ratio, MIN_BUY_AMOUNT)，再封顶到余额
   （余额 = 富途 API 账户可用资金，买入前强制实时刷新，不采用任何配置文件里的资金字段）
5. 下单：
   demo → 打印"拟投入金额"，返回模拟成交（价格=现价，数量=金额/价格）
   live → 按参考价折算 qty=int(金额/价格)，市价单
          （参考价取当前 tick 价；取不到则用行情快照 last_price）
6. 成交后状态写入：
   has_position=True、成本/数量/金额、buy_time/buy_reason、
   pair_type=pair_name、pair_sell_line=卖点、pair_sell_line_price、
   is_paired=True、check_level=T3、peak_price=成交价、stop_price=0、
   重置买入信号、清空错误重试计数
7. 写 positions.csv / pair_status.csv / tradeRecord.csv（BUY 行）
8. 刷新持仓缓存、更新 funds 快照、从买入队列移除
9. 钉钉推送「💰 买入」
```

**失败处理**：下单失败（返回 error）→ 进入 `_handle_unknown_error` 重试机制（见第 11 节）。

---

## 8. 卖出逻辑详解

### 8.1 三种卖出触发（按 tick 内优先级）

| 优先级 | 触发 | 判定 |
|---|---|---|
| 1 | 动态止盈/≥20%清仓 | 仅在 T0（盈利≥5%）时评估 `_check_take_profit` |
| 2 | 配对卖点 | `price ≥ pair_sell_line_price`（任何检查级别均生效） |
| — | （手动/外部） | 无，程序不监听外部指令 |

### 8.2 动态止盈算法（_check_take_profit）

```
profit_pct = (现价 - 成本) / 成本
现价 > 峰值 → 峰值 = 现价
stop_ratio = get_stop_loss(profit_pct)   # 按当前盈利率查表
├─ profit_pct ≥ 20% → 无条件立即清仓（v1.1：不再要求 stop_ratio==0）
└─ new_stop = 峰值 × (1 - stop_ratio)
   仅当 new_stop > 旧止损线时更新（止损线只升不降，ratchet 机制）
首次盈利 ≥5% → profit_triggered=True（启动跟踪）
profit_triggered 且 现价 ≤ 止损线 → 止盈清仓
```

**止盈回撤表锚点**（完整表 152 项，从 0.065→0.015 递减至 0.216→0；比例 = 允许从**峰值**回撤的幅度）：

| 当前盈利率 | 回撤比例 | 举例（峰值=105/110/115/120，成本≈100） |
|---|---|---|
| 5% ~ 6.5% | 1.5%（兜底值） | 峰值105 → 跌破 103.43 卖出 |
| 10% | 1.15% | 峰值110 → 跌破 108.74 卖出 |
| 15% | 0.65% | 峰值115 → 跌破 114.25 卖出 |
| 20% | 0.15% | 峰值120 → 跌破 119.82 卖出 |
| **≥20%** | — | **立即清仓** |

> 关键特性：比例按**当前**盈利率查表，但作用对象是**历史峰值**，且止损线只上移不下移——利润越高，止损线被"顶"得越贴近峰值；价格回落后比例要求虽放宽，但止损线不会倒退，保护已到手的利润。

### 8.3 配对卖点

买入时确定 `pair_type`（回撤/回调/过渡/极限）与对应卖点 `pair_sell_line_price`。此后任何 tick 一旦 `现价 ≥ 卖点价` 立即卖出，**不要求盈利 ≥5%**，是低于 5% 利润区间唯一会触发的止盈路径。

### 8.4 _execute_sell 执行细节

```
1. 无持仓 → 跳过
2. 下单：demo → 模拟成交（proceeds=现价×数量）
        live → 市价单卖出 int(数量)，订单号从返回 DataFrame 提取
3. 成交后：
   利润 = 到手金额 - 持仓成本；利润率 = 利润/成本
   缓存配对信息（pair_type/pair_sell_line/pair_sell_line_price/买入时间）
   has_position=False、清空持仓字段、is_sold=True
   写入冷却：cooldown_until = now + cooldown_days（默认7天）
   买入队列状态置「静默期_排除」
   恢复配对记录（保留历史买卖点与盈亏）供复盘
4. 写 positions.csv（删除该标的）/ pair_status.csv / tradeRecord.csv（SELL 行）
5. 刷新持仓缓存、更新 funds 快照
6. 钉钉推送「💸 卖出」（含盈亏金额与百分比）
```

---

## 9. 冷却、黑名单与观察列表

### 9.1 静默期冷却

- 触发：任何一次卖出成交；
- 时长：`COOLDOWN_DAYS`（.env）→ Excel `cooldown_days` → 默认 7 天；
- 记录：`cooldown.csv`（inst_id / sell_time / cooldown_until）；
- 效果：冷却期内 `_can_buy` 拒绝买入、`update_price` 跳过买入信号逻辑；
- 到期后：无需人工干预，`update_price` 每 tick 调用 `_sync_buy_queue`，队列状态自动从「静默期_排除」翻转为「待买入」（v1.1 修复）。

### 9.2 黑名单（手工维护 blacklist.csv）

| 列 | 含义 |
|---|---|
| inst_id | 标的代码 |
| reason | 拉黑原因 |
| add_time | 添加时间 |

- 拦截点（v1.1 修复前黑名单仅加载未使用）：`_can_buy`、`update_price` 买入信号分支、`_sync_buy_queue`（状态置「黑名单_排除」）；
- 效果：黑名单标的不再产生买入信号、不加入待买入队列。

### 9.3 观察列表（observe_list.csv，程序可自动追加）

- 程序在「连续 2 次未知错误」时自动将标的加入观察列表（原因=错误摘要）；
- 同样拦截买入信号；可手工从 CSV 中删除以解除。

### 9.4 买入队列状态机（buy_queue.csv）

```
待买入 ──买入成交──▶ 已持仓_排除 ──卖出成交──▶ 静默期_排除 ──冷却到期──▶ 待买入
                        ▲
                        └──重启对账：本地无记录但 API 有持仓 → 直接置 已持仓_排除

任意状态 ──加入黑名单────▶ 黑名单_排除
任意状态 ──加入观察列表──▶ 观察列表_排除
```

---

## 10. 持久化体系（CSV 全字段）

所有 CSV 均以 `utf-8-sig` 写入（Excel 直接打开不乱码），结构变更时自动迁移旧数据（v1.1 保留该机制）。

### 10.1 tradeRecord.csv（成交流水，追加式）

`timestamp, inst_id, direction, price, qty, amount, remaining_funds, position_price, profit, profit_pct, reason`

- `direction`: BUY / SELL；`reason`: 含触发原因与订单号（如 `反弹1%买入 (最低 98.3) | ordId=...`）。

### 10.2 positions.csv（本地持仓状态）

`inst_id, position_price, position_qty, position_amount, buy_time, buy_reason, pair_type, pair_sell_line, peak_price, stop_price, profit_triggered, check_level, last_check_time, current_price, market_value, unrealized_pnl, unrealized_pnl_pct`

- 卖出后该行删除；`current_price/market_value/unrealized_pnl` 为最近一次 tick 的估值。

### 10.3 cooldown.csv（静默期）

`inst_id, sell_time, cooldown_until`

### 10.4 high_low.csv（历史高低点）

`inst_id, all_time_high, all_time_low, last_update`

### 10.5 pair_status.csv（配对买卖历史）

`inst_id, pair_type, is_paired, buy_price, buy_time, sell_line, sell_line_price, is_sold, sell_time, sell_reason, profit, profit_pct`

### 10.6 buy_queue.csv（买入队列）

`inst_id, status, ref_high, ref_low, add_time, last_check_time`

### 10.7 funds.csv（资金快照，查看用）

`remaining_funds, total_position_market_value, total_unrealized_pnl, total_realized_pnl, net_asset_value, last_update`

> ⚠️ 持仓市值/浮动盈亏/已实现盈亏三列当前固定写 `0`，该文件仅用于快速查看可用资金，非完整资产表。

### 10.8 blacklist.csv / observe_list.csv

`inst_id, reason, add_time`

### 10.9 run.log（运行日志）

- ERROR 全量记录；INFO 仅记录启动/配置/买卖/止盈触发/冷却开始等关键事件；WARN 仅记录资金不足/连接错误；
- 所有行同时输出到终端。

---

## 11. 异常处理与告警

### 11.1 下单失败重试（_handle_unknown_error）

```
第1次失败：设置 _retry_after = now + 120 秒，120 秒内不再尝试该标的
第2次失败（600 秒窗口内）：将该标的加入观察列表、移出买入队列（防止无限重试）
成功成交：清空失败计数与重试时间
```

### 11.2 余额不足冷却

- 触发：`_can_buy` / `_execute_buy` 检测到余额 < MIN_BUY_AMOUNT；
- 效果：该标的 300 秒内不再尝试买入（v1.1 修复，此前冷却字段从未赋值，形同虚设）。

### 11.3 钉钉推送

| 事件 | 内容 |
|---|---|
| 买入成交 | 💰 买入：标的/买点/价格/金额 |
| 卖出成交 | 💸 卖出：标的/原因/价格/盈亏 |
| 卖出失败 | ⚠️ 卖出失败：标的/原因/数量 |
| ERROR 日志 | 🚨 系统异常 + 消息（webhook 配置后） |

- 加签方式：`timestamp + "\n" + secret` 做 HMAC-SHA256，base64 后拼到 webhook URL；
- 推送失败只打印日志，不影响主流程。

### 11.4 进程级保护

- 连接失败、配置缺失、订阅失败 → 打印明确提示并退出（不空转）；
- Ctrl+C → 优雅关闭行情/交易上下文、落资金快照、写退出日志。

---

## 12. 已知限制与注意点

| # | 限制 | 说明 |
|---|---|---|
| 1 | 实盘成交为估算 | live 买入按参考价折算数量、成交价取现价近似；实际以 Futu 账户为准，**重启时自动按 API 持仓对账修正** |
| 2 | 港股整手/碎股 | 实盘数量折算或卖出 `int(qty)` 可能触发 Futu 手数校验失败，失败走重试/观察列表保护，不会崩溃 |
| 3 | 资金口径 | 下单金额 = **Futu API 账户可用余额** × buy_ratio，以平台实时数据为准；配置中已无任何资金字段（total_funds 于 v1.2 移除），防止配置资金与真实账户不一致造成错乱 |
| 4 | funds.csv 市值列为 0 | 仅为资金查看快照，非完整资产报表 |
| 5 | 行情依赖 TICK 推送 | 休市期间无 tick，程序保活但不产生任何交易动作 |
| 6 | 系统参数部分失效 | profit_trigger/stop_loss 等参数已硬编码，改 Excel 无效（见 3.2 表） |
| 7 | 本地状态与 API 的一致性 | 盘中程序不会持续轮询 API 持仓（30 秒缓存），极端丢单场景以重启对账兜底 |

---

## 13. v1.1 修复记录

| # | 级别 | 问题 | 修复 |
|---|---|---|---|
| 1 | 🔴致命 | `_check_take_profit` 存在损坏语法行，文件无法编译 | 修复为正确日志语句 |
| 2 | 🔴严重 | `get_stop_loss` 正序遍历，动态止盈表完全失效（永远 1.5%） | 改为 `reversed()` 倒序匹配 |
| 3 | 🔴严重 | tick 时间 `float("2026-09-07 10:30:00.123")` 必崩 | `strptime` 解析（兼容带/不带毫秒） |
| 4 | 🔴严重 | 实盘 `place_order` 用不存在的 `direction` 参数、`qty=0` 非法、订单号存整个 DataFrame、未指定 trd_env | `trd_side`/`trd_env=REAL`、参考价折算数量、`_extract_order_id` |
| 5 | 🟠中 | 黑名单加载后从未使用，黑名单标的照常交易 | 三处拦截：`_can_buy`/`update_price`/`_sync_buy_queue` |
| 6 | 🟠中 | 余额不足冷却字段从未赋值 | `_can_buy`/`_execute_buy` 启用 300 秒冷却 |
| 7 | 🟠中 | `_execute_buy` 硬编码 110 与 MIN_BUY_AMOUNT 不一致 | 改用 `self.min_buy_amount` |
| 8 | 🟠中 | 成交数量可能为 0（除零/零持仓死锁） | `qty <= 0` 保护 |
| 9 | 🟡低 | Excel ref_high/ref_low 空值崩溃 | try/continue 容错 |
| 10 | 🟡低 | high_low/pair_status 空值转 float 崩溃 | `or 0` 容错 |
| 11 | 🟡低 | 止盈表 0.215 档位 Excel 浮点残留 9.02e-17 | 清理为 0.0 |
| 12 | 🟡低 | get_acc_balance 列名随版本变化 | avail_balance/avail_withdraw_cash 兼容 |
| 13 | 🟡低 | 买入队列每次同步无条件重写 CSV、静默期到期状态不刷新 | 状态变化才写盘 + 每 tick 同步 |
| 14 | ⚙️调整 | 立即清仓条件 `stop_ratio==0 and ≥20%`（实际 21.5% 才触发） | **`profit_pct ≥ 20%` 无条件立即清仓**（行为与文案对齐） |

### v1.2 变更记录

| # | 类型 | 变更 | 原因 |
|---|---|---|---|
| 15 | 🧹清理 | **移除 `total_funds` 全部引用**：`load_config` 默认参数删除该键（Excel 配置中即使写入也不再读取）、`SymbolState.__init__` 删除 `self.total_funds` 赋值 | 资金口径以富途 API 账户可用余额为准，配置资金字段与真实账户不一致会造成资金错乱 |

### v1.3 变更记录

| # | 类型 | 变更 | 原因 |
|---|---|---|---|
| 16 | 🛡️健壮性 | 新增 `_env_int`/`_env_float` 容错解析：`FUTU_OPEND_PORT`、`MIN_BUY_AMOUNT`、`COOLDOWN_DAYS` 自动清理复制粘贴残留字符（尾部 `\|`、空格、引号），无效值回退默认并打印告警 | 服务器 .env 中 `FUTU_OPEND_PORT=11111\|` 导致 `int()` 崩溃，启动即挂 |

---

*文档版本：v1.3（2026-09-07）· 与 futu_real.py v1.3 同步*
