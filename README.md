# 美团经营宝自动化助手

自动化获取美团/大众点评经营数据报表，减少手动操作，提升运营效率。

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-green.svg)

---

## 功能特性

- 🤖 **自动化数据采集** - 通过 Playwright 拦截网络请求获取 Cookie 和 mtgsig 凭证，调用美团 API
- 📊 **多板块支持** - 覆盖推广通、全站推广、简版看板、客流分析等核心板块
- 📁 **汇总报表生成** - 自动生成格式化报表，支持多门店合并（4个Sheet一页览）
- 🔒 **凭证管理** - 安全存储登录态，支持一键刷新
- 🌐 **Web 界面** - 提供 API 接口和可视化页面，方便集成和调用

---

## 支持的板块

| 板块 | 状态 | 说明 |
|------|------|------|
| 推广通 | ✅ 已完成 | 广告花费、曝光、点击、转化、竞争分析 |
| 全站推广 | ✅ 已完成 | 全站推广数据 |
| 简版看板 | ✅ 已完成 | 流量数据、星级数据、门店评价 |
| 客流分析 | ✅ 已完成 | 客流统计、引流用户、在线咨询、门店交易、评价概览、星级概览 |
| 评论统计 | ✅ 已完成 | 新版 Playwright API |
| 星级评分 | ✅ 已完成 | 点评星级、美团星级 |
| 门店评价 | ✅ 已完成 | 用户评论、评分、商户回复 |

---

## 快速开始

### 环境要求

- Python 3.8+
- Windows / Linux 系统
- Chrome/Chromium 浏览器

### 安装依赖

```bash
pip install -r requirements.txt
playwright install chromium
```

### 获取凭证

```bash
python common/credentials.py
```

自动打开浏览器进行扫码登录，登录成功后自动保存凭证。

### 启动 Web 服务

```bash
python -m web.api
```

访问地址：http://127.0.0.1:5000

### 生成报表

**方式一：Web 界面**
1. 打开 http://127.0.0.1:5000
2. 选择门店、平台、日期范围
3. 点击生成报表

**方式二：API 调用**
```bash
curl -X POST http://127.0.0.1:5000/api/summary/generate \
  -H "Content-Type: application/json" \
  -d '{
    "门店列表": ["1933643130"],
    "平台": 0,
    "评论统计平台": 1,
    "评论统计日期范围": 2,
    "日期范围": {"begin": "2026-05-01", "end": "2026-05-20"}
  }'
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
│   └── credentials.py        # 凭证获取工具
│
├── tuiguangtong/             # 推广通板块
│   ├── client.py             # API 客户端
│   └── excel_generator.py    # Excel 生成器
│
├── quantianzhan/             # 全站推广板块
├── simple_board/             # 简版看板板块
├── customer_flow/            # 客流分析板块
├── review_statistics/        # 评论统计板块（Playwright API）
├── star_rating/             # 星级评分板块
├── shop_review/             # 门店评价板块
│
├── web/                      # Web 服务
│   ├── api.py               # Flask API
│   └── summary.html         # 前端页面
│
├── config/                   # 配置文件
│   └── reports/             # 报表输出目录
│
├── docs/                     # 文档
│   └── 部署方案.md          # 服务器部署指南
│
├── summary_report_generator.py  # 汇总报表生成器
└── requirements.txt            # 依赖列表
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

## API 接口

### 生成汇总报表

```
POST /api/summary/generate
```

**请求体：**
```json
{
  "门店列表": ["1933643130"],
  "平台": 0,
  "评论统计平台": 1,
  "评论统计日期范围": 2,
  "日期范围": {
    "begin": "2026-02-01",
    "end": "2026-03-31"
  }
}
```

**响应：** Excel 文件下载

### 获取门店列表

```
GET /api/shops?page=1&page_size=100
```

### 更新门店数据

```
POST /api/shops/refresh
```

---

## 常见问题

**Q: 请求返回 401/403 Unauthorized**
> 凭证已过期，运行 `python common/credentials.py --force` 重新获取

**Q: Excel 文件被占用**
> 关闭 Excel 后重试

**Q: 报表生成失败，显示"数据获取失败"**
> 检查是否有多余的 Python 进程占用端口，先执行 `taskkill //F //IM python.exe` 终止所有进程后重新启动

**Q: Playwright 浏览器启动失败**
> 运行 `playwright install chromium` 重新安装浏览器

---

## 服务器部署

详细部署指南请参考 [部署方案](docs/部署方案.md)

---

## 安全注意

- 凭证文件包含登录态，请勿上传 `auth/` 目录到公开仓库
- 仅供个人/内部数据备份使用
- 避免过于频繁的请求，可能触发风控

---

## License

MIT License
