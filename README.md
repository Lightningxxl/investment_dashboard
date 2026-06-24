# Investment Dashboard

中国50研选池子基金组合工作台。当前版本是纯静态页面，SQLite 是主数据层，前端读取从 SQLite 生成的 `data/dashboard_data.js` 快照。

## 本地运行

```bash
python3 -m http.server 8765
```

然后打开 `http://127.0.0.1:8765/`。

## 文件结构

- `index.html`: dashboard 主页面
- `lucide-lite.js`: 本地轻量图标运行时
- `data/investment_dashboard.sqlite`: 主数据层，保存 PDF 全池、推荐标的、匹配关系、净值/回撤/规模/指标序列
- `data/dashboard_data.js`: 前端唯一读取的数据快照
- `data/dashboard_data.json`: 同一份前端快照，便于检查
- `data/source_china50_fund_pool.json`: PDF 全池结构化源数据
- `data/china50_fund_pool.csv`: PDF 全池 CSV 导出
- `scripts/build_data_store.py`: 从源 JSON 和推荐标的 zip 生成 SQLite + 前端快照
- `deploy/ff2-sync.sh`: ff2 服务器通过 git 同步的辅助脚本

## 重建数据层

```bash
python3 scripts/build_data_store.py ~/Downloads/中国50研选池推荐标的整理.zip
```

脚本需要 `openpyxl`。它会读取每个 Excel 的净值走势、收益走势、动态回撤、区间收益、收益指标、风险指标和产品规模，写入 `data/investment_dashboard.sqlite`，再生成 `data/dashboard_data.json/js` 给前端使用。高置信匹配到 PDF 产品名的标的会进入 dashboard 模型；需要人工确认或未匹配的标的保留在 SQLite 和 JSON 里。

## 当前模型口径

- 左侧第 1 步定义投资政策：预期收益下限、最大回撤预算、波动预算、Beta 暴露和组合敞口集中度。
- 左侧第 2 步定义参数偏好：历史样本长度、低大盘相关性、低集中度风险、数据口径质量按 0-5 重要性加权为参数质量分。
- 左侧参数采用草稿/应用机制：调整参数后点击“应用并计算”才重新生成组合，右侧列表搜索和排序只刷新表格。
- 高置信 Excel 明细匹配成功时，收益、波动、历史回撤、净值跨度和产品规模优先使用 Excel 数据。
- 预期收益 μ 使用历史年化收益乘以收益置信折扣；收益置信折扣由参数质量分决定。
- 等权模式固定每只基金权重为 `1 / K`，先构建可遍历资产池，再精确枚举所有 K 只基金组合。
- 质量倾斜模式先构建可投标的池和优化资产池，再用固定种子的确定性多起点搜索生成风险-收益可行前沿。
- 未匹配产品的 Beta、相关性、波动率基于 PDF 中年度收益估算，属于年度估算口径。

## ff2 Git 同步

首次部署：

```bash
ssh ff2
mkdir -p /home/xavierx/www
git clone https://github.com/Lightningxxl/investment_dashboard.git /home/xavierx/www/zy-sj-dashboard
cd /home/xavierx/www/zy-sj-dashboard
python3 -m http.server 20242 --bind 127.0.0.1
```

后续同步：

```bash
ssh ff2
cd /home/xavierx/www/zy-sj-dashboard
git pull --ff-only origin main
```

当前 ff2 的 `zy-sj-dashboard` user systemd 服务可以直接沿用这个目录。
