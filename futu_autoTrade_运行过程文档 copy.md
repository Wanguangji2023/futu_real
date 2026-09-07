  1. 配对卖点（固定止盈位，如「回撤上沿」）；===》取消
  OpenTradeContext 交易（账户资金/持仓/下单，demo/live）===》改为由.env中设置实盘还是模拟，分市场设置
  本程序.env中主要设哪些参考要给出，直接给完整的.env文件


**Sheet 2：`系统参数`**（key / value 两列）====》

| 参数 | 默认值 | 实际生效情况 |
|---|---|---|
| `total_funds` | 1000000 | ⚠️ 载入 `SymbolState.total_funds`，但**下单金额实际按账户余额×buy_ratio 计算，此参数未参与下单** |
| `buy_ratio` | 0.01 | ✅ 单笔买入 = 账户余额 × 比例 |
| `cooldown_days` | 7 | ✅ 卖出后静默天数（.env 的 COOLDOWN_DAYS 可覆盖） |
| `profit_trigger` | 0.05 | ❌ 硬编码 0.05（T0 门槛），参数不生效 |
| `stop_loss` | 0.015 | ❌ 硬编码（get_stop_loss 兜底值），参数不生效 |
| `profit_step` / `break_alert_max` / `break_alert_interval` / `push_interval` / `data_source` / `trade_mode` | — | ❌ 保留字段，当前逻辑未使用 |


**买卖点配对**（每对 = 同区域的「2下」买点 + 「上」卖点）：这部分不能保证5%收益

| 买点（跌破激活） | 对应卖点（升破卖出） | 名称 |
|---|---|---|
| pullback2_down | pullback_up | 回撤 |
| callback2_down | callback_up | 回调 |
| transition2_down | transition_up | 过渡 |
| limit2_down | limit_up | 极限 |