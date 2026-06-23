# Investment Dashboard

中国50研选池子基金组合工作台。当前版本是纯静态页面，数据来自 `中国50研选池完整版_2026-06-18.pdf` 的结构化抽取结果。

## 本地运行

```bash
python3 -m http.server 8765
```

然后打开 `http://127.0.0.1:8765/`。

## 文件结构

- `index.html`: dashboard 主页面
- `china50_fund_pool_data.js`: 页面直接读取的数据源
- `china50_fund_pool_data.json`: 同一份数据的 JSON 版本
- `china50_fund_pool_data.csv`: 表格导出
- `lucide-lite.js`: 本地轻量图标运行时
- `deploy/ff2-sync.sh`: ff2 服务器通过 git 同步的辅助脚本

## 当前模型口径

- 产品评分由历史样本长度、低大盘相关性、低集中度风险、数据口径质量四个维度按 0-5 重要性加权得到。
- 调整后收益使用历史收益乘以可信度折扣；可信度由综合评分和数据口径质量共同决定。
- 组合推荐在当前目标约束、评分偏好、产品权重上限和策略大类上限下做模拟筛选。
- Beta、相关性、波动率基于 PDF 中年度收益估算，属于年度 proxy。

## ff2 Git 同步

首次部署：

```bash
ssh ff2
mkdir -p /home/xavierx/www
git clone https://github.com/Lightningxxl/investment_dashboard.git /home/xavierx/www/investment_dashboard
cd /home/xavierx/www/investment_dashboard
python3 -m http.server 20242 --bind 127.0.0.1
```

后续同步：

```bash
ssh ff2
cd /home/xavierx/www/investment_dashboard
git pull --ff-only origin main
```

如果继续沿用现有 `zy-sj-dashboard` systemd 服务，把服务的 `WorkingDirectory` 和 `--directory` 指到 git clone 出来的目录即可。
