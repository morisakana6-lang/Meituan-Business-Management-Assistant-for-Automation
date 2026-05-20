# 门店评论板块

## 板块说明

门店评论数据获取，对应目录 `shop_review/`。

## API 信息

- **URL**: `https://e.dianping.com/gateway/merchant/review/list`
- **方法**: POST
- **Content-Type**: application/json

## Platform 参数说明

**重要**：本模块的 platform 参数含义与项目通用规则不同：

| platform 值 | 含义 | 说明 |
|-------------|------|------|
| `0` | 不支持 | 会返回"平台不支持"错误 |
| `1` | 点评 | 获取点评平台的评论 |
| `2` | 美团 | 获取美团平台的评论 |

### 项目通用规则（供参考其他模块）

| platform 值 | 含义 |
|-------------|------|
| `0` | 全平台（点评+美团） |
| `1` | 点评 |
| `2` | 美团 |

**注意**：如果需要同时获取点评和美团的评论，需要分别调用两次 API（platform=1 和 platform=2），然后合并结果。

## 配置文件

`shop_config.json` 示例：

```json
{
  "板块名称": "门店评论",
  "search_key": ["1933643130", "57904793"],
  "platform": 1,
  "日期范围": {
    "begin": "2026-05-01",
    "end": "2026-05-18"
  }
}
```

## 使用方法

### 单门店查询

```python
from shop_review.client import ShopReviewClient

client = ShopReviewClient()
reviews = client.get_reviews(
    shop_id="1933643130",
    platform="1",  # 1=点评，2=美团
    begin_date="2026-05-01",
    end_date="2026-05-18"
)
```

### 多门店查询

```python
from shop_review.client import ShopReviewClient

client = ShopReviewClient()
shop_ids = ["1933643130", "57904793"]

for shop_id in shop_ids:
    reviews = client.get_reviews(
        shop_id=shop_id,
        platform="1",
        begin_date="2026-05-01",
        end_date="2026-05-18"
    )
    print(f"门店 {shop_id}: {len(reviews)} 条评论")
```

## 指标说明

本模块获取的是**评论明细数据**，不是汇总指标数据。

返回的字段包括：
- `shop_id`: 门店ID
- `shop_name`: 门店名称
- `review_id`: 评论ID
- `content`: 评论内容
- `star`: 评分（1-5星）
- `score_detail`: 细致评分（效果、服务、环境）
- `pictures`: 评论图片
- `create_time`: 评论时间
- `reply_content`: 回复内容
- `reply_time`: 回复时间

## 测试

```bash
# 运行多门店测试
python shop_review/test_multi_shop.py

# 查看帮助信息
python shop_review/client.py
```
