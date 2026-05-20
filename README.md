# 美团经营宝自动化助手

自动化获取美团/大众点评经营数据报表，减少手动操作，提升运营效率。

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-green.svg)

---

## 功能特性

- 🤖 **自动化数据采集** - 通过 Playwright 拦截网络请求获取 Cookie 和 mtgsig 凭证，调用美团 API
- 📊 **多板块支持** - 覆盖推广通、评论统计、星级评分、门店评价、在线咨询、交易分析、客流统计等
- 📁 **Excel 报表生成** - 自动生成格式化报表，支持多门店合并
- 🔒 **凭证管理** - 安全存储登录态，支持一键刷新

---

## 支持的板块

| 板块 | 状态 | 说明 |
|------|------|------|
| 推广通 | ✅ 已完成 | 广告花费、曝光、点击、转化等 |
| 评论统计 | ✅ 已完成 | 累计评价、新增评价、好评/差评分析 |
| 星级评分 | ✅ 已完成 | 点评星级、美团星级 |
| 门店评价 | ✅ 已完成 | 用户评论、评分、商户回复 |
| 在线咨询 | ✅ 已完成 | 咨询人数、留资转化 |
| 交易分析 | ✅ 已完成 | 交易金额、订单量 |
| 客流统计 | ✅ 已完成 | 客流量分析 |
| 全站推广 | 🔄 待完善 | - |
| 简单面板 | 🔄 待完善 | - |

---

## 快速开始

### 环境要求

- Python 3.8+
- Windows 系统
- Chrome 浏览器

### 安装依赖

```bash
pip install requests playwright openpyxl
python -m playwright install chromium
```

### 获取凭证

```bash
python common/credentials.py
```

自动打开浏览器进行扫码登录，登录成功后自动保存凭证。

### 生成报表

```bash
# 推广通报表
python tuiguangtong/excel_generator.py

# 评论统计报表
python review_statistics/excel_generator.py

# 星级评分报表
python star_rating/excel_generator.py
```

---

## 项目结构

```
meituan/
├── auth/                      # 凭证目录（不提交到版本控制）
│   ├── auth.json             # Playwright 登录态
│   └── credentials.json      # Cookie + mtgsig
│
├── common/                    # 公共模块
│   └── credentials.py         # 凭证获取工具
│
├── tuiguangtong/              # 推广通
├── review_statistics/          # 评论统计
├── star_rating/               # 星级评分
├── shop_review/               # 门店评价
├── online_consultation/        # 在线咨询
├── trade_analysis/            # 交易分析
├── customer_flow/             # 客流统计
├── quantianzhan/             # 全站推广
├── simple_board/             # 简单面板
└── legacy/                   # 旧版本参考
```

---

## 配置说明

每个板块目录下有 `shop_config.json` 配置文件：

```json
{
  "板块名称": "星级评分",
  "search_key": ["1933643130"],
  "platform": 0,
  "日期范围": {
    "begin": "2026-04-01",
    "end": "2026-04-30"
  }
}
```

| 参数 | 说明 |
|------|------|
| search_key | 门店 ID，支持多门店 |
| platform | 0=全平台，1=点评，2=美团 |
| 日期范围 | 格式 YYYY-MM-DD |

---

## API 端点

| 板块 | URL | 方法 |
|------|-----|------|
| 推广通 | `e.dianping.com/shopdiy/report/...` | GET |
| 评论统计 | `e.dianping.com/gateway/adviser/data` | GET |
| 星级评分 | `e.dianping.com/gateway/adviser/data` | POST |
| 在线咨询 | `e.dianping.com/mda/v5/onlineConsultant` | POST |

---

## 常见问题

**Q: 请求返回 401 Unauthorized**
> 凭证已过期，运行 `python common/credentials.py --force` 重新获取

**Q: Excel 文件被占用**
> 关闭 Excel 后重试

---

## 安全注意

- 凭证文件包含登录态，请勿上传 `auth/` 目录到公开仓库
- 仅供个人/内部数据备份使用
- 避免过于频繁的请求，可能触发风控

---

## License

MIT License
