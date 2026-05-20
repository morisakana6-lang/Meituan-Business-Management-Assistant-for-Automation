"""
门店评论数据获取客户端

=== 板块归属 ===
门店评论板块，对应目录 shop_review/

=== API 信息 ===
URL: https://e.dianping.com/gateway/merchant/review/list
方法: POST
Content-Type: application/json

=== Payload 参数 ===
- platform: 平台（0=全平台，1=点评，2=美团）
- shopIds: 门店ID数组
- startTime: 开始时间戳（毫秒）
- endTime: 结束时间戳（毫秒）
- pageNo: 页码
- pageSize: 每页条数（最大20）
- aiReply: 是否AI回复

=== 响应字段 ===
- reviewDetail.reviewInfo.reviewId: 评论ID
- reviewDetail.reviewInfo.content: 评论内容
- reviewDetail.reviewInfo.star: 评分
- reviewDetail.reviewInfo.reviewPicList: 图片列表
- reviewDetail.reviewInfo.addTime: 评论发布时间
- reviewDetail.replyList: 回复列表

=== 使用方法 ===
from shop_review.client import ShopReviewClient

client = ShopReviewClient()
reviews = client.get_reviews(
    shop_id="1933643130",
    begin_date="2026-05-01",
    end_date="2026-05-16"
)
"""

import requests
import json
import time
from datetime import datetime, timedelta
from typing import Dict, List

import sys
import os

# 获取项目根目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)


def load_credentials():
    """直接从凭证文件读取"""
    cred_file = os.path.join(_PROJECT_ROOT, 'auth', 'credentials.json')
    if os.path.exists(cred_file):
        with open(cred_file, encoding='utf-8') as f:
            data = json.load(f)
            return data.get('cookie', ''), data.get('mtgsig', '')
    return '', ''


def load_all_shop_ids():
    """从 all_shop_ids.json 加载所有门店ID"""
    shop_ids_file = os.path.join(_CURRENT_DIR, 'all_shop_ids.json')
    if os.path.exists(shop_ids_file):
        with open(shop_ids_file, encoding='utf-8') as f:
            data = json.load(f)
            return data.get('shop_ids', [])
    return []


class ShopReviewClient:
    """门店评论数据客户端"""

    BASE_URL = "https://e.dianping.com/gateway/merchant/review/list"

    DEFAULT_HEADERS = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Content-Type": "application/json",
        "Origin": "https://e.dianping.com",
        "Referer": "https://e.dianping.com/vg-platform-reviewmanage/rating-management/index.html",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def _build_url(self, mtgsig: str = None) -> str:
        """构建带mtgsig参数的URL"""
        sig = mtgsig or self.mtgsig
        if not sig:
            raise ValueError("mtgsig is required")
        return f"{self.BASE_URL}?yodaReady=h5&csecplatform=4&csecversion=4.2.0&mtgsig={sig}"

    def _date_to_timestamp_ms(self, date_str: str) -> int:
        """日期字符串转时间戳（毫秒），日期当天00:00:00"""
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return int(dt.timestamp() * 1000)

    def get_reviews(
        self,
        shop_id: str,
        platform: str = "1",
        begin_date: str = None,
        end_date: str = None,
        page_size: int = 20,
    ) -> List[Dict]:
        """
        获取门店评论列表

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            page_size: 每页条数（最大20）

        Returns:
            评论列表（原始数据结构）
        """
        # 设置默认日期
        if not end_date:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if not begin_date:
            begin_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        # 结束日期加一天再减1秒（确保包含当天）
        end_dt = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
        end_ts = int(end_dt.timestamp() * 1000) - 1

        start_ts = self._date_to_timestamp_ms(begin_date)

        payload = {
            "platform": int(platform),
            "shopIds": [shop_id],
            "startTime": start_ts,
            "endTime": end_ts,
            "aiReply": False,
            "pageNo": 1,
            "pageSize": page_size,
        }

        url = self._build_url()
        self.session.headers["Cookie"] = self.cookie

        try:
            response = self.session.post(url, json=payload, timeout=30)
            if response.status_code != 200:
                print(f"请求失败 (状态码: {response.status_code})")
                return []

            result = response.json()
            if result.get("code") != 200:
                print(f"API错误: {result.get('msg')}")
                return []

            data = result.get("data", {})
            all_reviews = data.get("reviewDetails", [])
            total_count = data.get("totalSize", 0)

            print(f"获取{len(all_reviews)}条评论，总{total_count}条")
            return all_reviews

        except Exception as e:
            print(f"请求异常: {e}")
            return []

    def get_reviews_all_pages(
        self,
        shop_ids: List[str],
        platform: str = "1",
        begin_date: str = None,
        end_date: str = None,
        page_size: int = 20,
        batch_size: int = 50,
    ) -> List[Dict]:
        """
        获取多个门店的所有评论（自动分页+分批）

        Args:
            shop_ids: 门店ID列表
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            page_size: 每页条数（最大20）
            batch_size: 每批门店数量（默认50）

        Returns:
            所有页合并后的评论列表
        """
        # 设置默认日期
        if not end_date:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if not begin_date:
            begin_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        # 结束日期加一天再减1秒（确保包含当天）
        end_dt = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
        end_ts = int(end_dt.timestamp() * 1000) - 1

        start_ts = self._date_to_timestamp_ms(begin_date)

        url = self._build_url()
        self.session.headers["Cookie"] = self.cookie

        all_reviews = []

        print(f"开始获取 {len(shop_ids)} 个门店的评论（分批，每批 {batch_size} 个）...")

        # 分批处理门店
        total_batches = (len(shop_ids) + batch_size - 1) // batch_size
        for batch_idx in range(total_batches):
            batch_start = batch_idx * batch_size
            batch_end = min(batch_start + batch_size, len(shop_ids))
            current_batch = shop_ids[batch_start:batch_end]

            print(f"\n批次 {batch_idx + 1}/{total_batches}: 处理门店 {batch_start + 1} ~ {batch_end}")

            page_no = 1
            batch_has_data = False

            while True:
                payload = {
                    "platform": int(platform),
                    "shopIds": current_batch,
                    "startTime": start_ts,
                    "endTime": end_ts,
                    "aiReply": False,
                    "pageNo": page_no,
                    "pageSize": page_size,
                }

                try:
                    response = self.session.post(url, json=payload, timeout=60)
                    if response.status_code != 200:
                        print(f"  请求失败 (状态码: {response.status_code})")
                        break

                    result = response.json()
                    if result.get("code") != 200:
                        # 如果是权限问题，跳过这批
                        if "不在该账号管辖范围内" in str(result.get("msg", "")):
                            print(f"  跳过无权限门店")
                            break
                        print(f"  API错误: {result.get('msg')}")
                        break

                    data = result.get("data", {})
                    reviews = data.get("reviewDetails", [])
                    total_size = data.get("totalSize", 0)

                    if page_no == 1:
                        if total_size > 0:
                            print(f"  总评论数: {total_size}")
                            batch_has_data = True

                    all_reviews.extend(reviews)

                    # 判断是否还有更多数据
                    if not reviews or len(reviews) < page_size:
                        break

                    page_no += 1
                    time.sleep(0.3)  # 避免请求过快

                except Exception as e:
                    print(f"  请求异常: {e}")
                    break

            if batch_has_data:
                print(f"  本批获取到 {len(all_reviews)} 条评论（累计）")

        print(f"\n共获取到 {len(all_reviews)} 条评论")
        return all_reviews

    def get_review_detail(self, review: Dict) -> Dict:
        """
        解析单条评论详情

        Args:
            review: 原始评论数据（reviewDetails中的项）

        Returns:
            格式化后的评论数据
        """
        review_info = review.get("reviewDetail", {}).get("reviewInfo", {})

        # 解析门店信息
        shop_info = review.get("shopInfo", {})
        shop_id = shop_info.get("shopId", "")
        shop_name = shop_info.get("shopName", "")

        # 解析评论ID
        review_id = review_info.get("reviewId", "")

        # 解析评论内容
        content = review_info.get("content", "")

        # 解析评分（star是50表示5星）
        star = review_info.get("star", "")
        if star:
            star = str(int(star) // 10)  # 50 -> 5

        # 解析细致评分（效果、服务、环境）
        score_list = review_info.get("scoreList") or []
        score_detail = ""
        for score_item in score_list:
            title = score_item.get("title", "")
            score = score_item.get("score", "")
            if score_detail:
                score_detail += f"\n{title}：{score}分"
            else:
                score_detail = f"{title}：{score}分"

        # 解析用户图片（换行分隔）
        pic_list = review_info.get("reviewPicList") or []
        picture_urls = "\n".join([p.get("url", "") for p in pic_list if p.get("url")])

        # 解析评论时间
        add_time = review_info.get("addTime", "")
        create_time = ""
        if add_time:
            try:
                ts = int(str(add_time)[:10])
                create_time = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
            except:
                pass

        # 解析商户回复（支持多条回复）
        reply_list = review.get("reviewDetail", {}).get("replyList") or []
        reply_contents = []
        reply_times = []
        for idx, reply in enumerate(reply_list, 1):
            reply_body = reply.get("replyBody", "")
            reply_add_time = reply.get("addTime", "")
            if reply_body:
                # 用数字标记区分多条回复
                reply_contents.append(f"回复{idx}：{reply_body}")
            if reply_add_time:
                try:
                    ts = int(str(reply_add_time)[:10])
                    reply_times.append(datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S"))
                except:
                    pass

        # 多条回复用换行符合并
        reply_content = "\n".join(reply_contents)
        reply_time = ", ".join(reply_times)

        return {
            "shop_id": str(shop_id),
            "shop_name": shop_name,
            "review_id": str(review_id),
            "content": content,
            "star": star,
            "score_detail": score_detail,
            "pictures": picture_urls,
            "create_time": create_time,
            "reply_content": reply_content,
            "reply_time": reply_time,
        }


if __name__ == "__main__":
    """测试门店评论API"""
    print("=" * 50)
    print("门店评论 API 测试")
    print("=" * 50)

    config_file = os.path.join(_CURRENT_DIR, 'shop_config.json')
    if os.path.exists(config_file):
        with open(config_file, encoding='utf-8') as f:
            config = json.load(f)
            print(f"\n配置文件: {config}")

    shop_id = config.get("search_key", "1933643130")
    # 支持数组或逗号分隔的字符串
    if isinstance(shop_id, list):
        shop_ids = shop_id
    elif isinstance(shop_id, str) and "," in shop_id:
        shop_ids = [s.strip() for s in shop_id.split(",")]
    else:
        shop_ids = [shop_id]
    shop_id = shop_ids[0]  # 主要用于文件名
    platform = config.get("platform", 1)
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")

    print(f"\n使用参数:")
    print(f"  门店ID: {shop_ids}")
    print(f"  平台: {platform}")
    print(f"  日期: {begin_date} 至 {end_date}")

    client = ShopReviewClient()
    reviews = client.get_reviews(shop_id, str(platform), begin_date, end_date)

    print(f"\n获取到 {len(reviews)} 条评论")
    if reviews:
        detail = client.get_review_detail(reviews[0])
        print(f"\n第一条评论详情:")
        for k, v in detail.items():
            print(f"  {k}: {v}")