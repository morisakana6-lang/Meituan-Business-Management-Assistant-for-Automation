# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

美团经营宝报表自动化工具：通过 Playwright 拦截网络请求获取 Cookie 和 mtgsig 凭证，调用美团 API 获取各类经营数据，生成 Excel 报表。

## 常用命令

```bash
# 获取/更新凭证（首次运行或凭证过期时）
python common/credentials.py
python common/credentials.py --force  # 强制重新登录

# 运行各板块测试脚本（以星级估分为例）
python star_rating/test_client.py

# 直接生成Excel报表（以星级评分为例）
python star_rating/excel_generator.py
```

## 目录结构

```
meituan/
├── auth/                    # 凭证目录（不提交到版本控制）
│   ├── auth.json           # Playwright 登录态持久化
│   └── credentials.json    # Cookie + mtgsig
│
├── common/                 # 公共模块
│   └── credentials.py     # Playwright 凭证获取
│
├── tuiguangtong/           # 推广通板块
│   ├── shop_config.json    # 门店配置
│   ├── shop_mapping.json   # 门店ID映射表
│   ├── client.py           # API 客户端
│   └── excel_generator.py  # Excel 生成器
│
├── review_statistics/       # 评论统计板块
│   ├── shop_config.json    # 门店配置
│   ├── client.py           # API 客户端（含get_shop_name）
│   └── excel_generator.py  # Excel 生成器
│
├── star_rating/            # 星级评分板块
│   ├── shop_config.json    # 门店配置
│   ├── client.py           # API 客户端（含get_shop_name）
│   └── excel_generator.py  # Excel 生成器
│
├── shop_review/            # 门店评价板块
├── online_consultation/    # 在线咨询板块
├── trade_analysis/         # 交易分析板块
├── customer_flow/          # 客流统计板块
├── quantianzhan/           # 全站推广板块
└── simple_board/           # 简单面板板块
```

## 板块模块规范

每个业务板块采用统一架构：

| 文件 | 作用 |
|------|------|
| `shop_config.json` | 板块配置：search_key(门店ID)、platform、日期范围 |
| `client.py` | API 客户端：凭证管理、请求发送、响应解析 |
| `excel_generator.py` | Excel 生成器：样式定义、数据写入、文件保存 |
| `reports/` | Excel 报表输出目录 |

### shop_config.json 格式

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

### get_shop_name 函数

`review_statistics` 和 `star_rating` 模块实现了从 `tuiguangtong/shop_mapping.json` 获取门店名称：

```python
def load_shop_mapping():
    mapping_file = os.path.join(_PROJECT_ROOT, 'tuiguangtong', 'shop_mapping.json')
    # 返回 {shop_id: shop_name} 映射

def get_shop_name(shop_id: str) -> str:
    mapping = load_shop_mapping()
    return mapping.get(str(shop_id), '')
```

## 凭证体系

- `auth.json`：Playwright 登录态，可复用数周
- `credentials.json`：Cookie + mtgsig，会过期
- mtgsig 通过 Playwright 拦截请求从 URL 正则提取

**Cookie 说明**：收集 `dianping.com` 和 `meituan.com` 域名的 Cookie。`JSESSIONID` 可能获取不到，但不影响请求（只需 `edper` 即可正常调用 API）。

## API 端点

| 板块 | URL | 方法 |
|------|-----|------|
| 推广通 | `https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport` | GET |
| 评论统计 | `https://e.dianping.com/gateway/adviser/data` | GET |
| 星级评分 | `https://e.dianping.com/gateway/adviser/data` | POST |
| 在线咨询 | `https://e.dianping.com/mda/v5/onlineConsultant` | POST |

**platform 参数值**：`platform=0` 代表点评+美团全平台，1代表点评，2代表美团。

## 数值格式

API 返回的数值可能包含逗号（如 `"146,504"`），必须先 `.replace(",", "")` 再转换为浮点数。

## 规则

### 回复风格
- 每次回复结尾要有精炼总结，压缩语言抓住关键
- 总结只说结论和关键信息，不解释过程
- 直接看diff能知道的不说

### 记忆触发时机
以下情况立即存入记忆，不要等对话结束：
1. 用户说"不对"、"你错了"、"不要"的时候
2. 用户说"没问题"、"对的"、"就这样"的时候
3. 用户提供了我做法的正确答案时

存入记忆后检查 MEMORY.md 是否需要同步。

### Excel 文件占用检查
生成 Excel 文件前，必须检查目标文件是否被占用（Excel 打开时会出现这个问题）。如果文件存在且被占用，提示用户关闭文件后再重试。

### 凭证过期处理
401/403 错误时提示用户运行 `python common/credentials.py --force`，不要频繁重新获取。

### 禁止修改凭证文件
**禁止以任何形式直接修改 `auth/credentials.json` 或 `auth/auth.json` 文件**。凭证更新必须通过 Playwright 流程进行。

## 依赖

Python 3.8+, requests, playwright, playwright-stealth, openpyxl

## Git / GitHub 操作规范

### Git 初始化与远程仓库连接

新项目首次上传 GitHub 步骤：

```bash
# 1. 初始化 git 仓库
git init

# 2. 添加 remote（使用 token 认证）
git remote add origin https://github.com/[username]/[repo].git

# 3. 设置 remote URL 包含 token（方便后续 push）
git remote set-url origin https://[token]@github.com/[username]/[repo].git
```

### 创建 .gitignore

新项目必须创建 `.gitignore`，排除以下内容：
- `auth/` 目录（凭证文件）
- `*.json`（凭证和配置）
- `__pycache__/`（Python 缓存）
- `*.pyc`（编译文件）
- `.playwright-mcp/`（Playwright 缓存）
- `reports/`、`*.xlsx`、`*.zip`（报表和压缩包）
- `.DS_Store`、`*.log`（系统文件）

### 提交与推送

```bash
# 添加所有文件
git add -A

# 提交（首次提交）
git commit -m "Initial commit: [项目名称]"

# 推送并设置上游分支
git push -u origin [branch]
```

### 分支管理

- GitHub 新仓库默认分支为 `main`，优先使用 `main` 作为主分支
- 本地分支名应与远程保持一致
- 如需重命名分支：
  ```bash
  git branch -m old_name new_name
  ```

### GitHub 认证方式

**Personal Access Token（推荐）**：
1. GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. 生成新 token，勾选 `repo` 权限
3. 将 token 嵌入 remote URL：`https://[token]@github.com/...`

### 禁止操作

- 严禁将 `auth/` 目录推送到远程仓库
- 严禁在 commit 中包含凭证文件
- 严禁 push --force 到 main/master 主分支（除非有充分理由并告知用户）
