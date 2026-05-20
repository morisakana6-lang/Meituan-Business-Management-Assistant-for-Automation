"""
交易分析 API 测试文件

基于 curl 测试验证的美团交易分析 API
"""

import requests
import json
import re
from datetime import datetime


class TradeAnalysisClient:
    """交易分析 API 客户端"""

    BASE_URL = "https://e.dianping.com/mda/v5/trade"

    # 请求头
    HEADERS = {
        "accept": "application/json, text/plain, */*",
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "zh-CN,zh;q=0.9",
        "connection": "keep-alive",
        "content-type": "application/x-www-form-urlencoded",
        "origin": "https://h5.dianping.com",
        "referer": "https://h5.dianping.com/",
        "sec-ch-ua": "\"Chromium\";v=\"148\", \"Google Chrome\";v=\"148\", \"Not/A)Brand\";v=\"99\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\"",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }

    # Cookie（从 auth/credentials.json 读取）
    COOKIE = "AWPTALOS32837=; WEBDFPID=1778731404479GQQAKEEfd79fef3d01d5e9aadc18ccd4d0c95073836-1778731404479-1778731404479GQQAKEEfd79fef3d01d5e9aadc18ccd4d0c95073836; _lxsdk_cuid=19e24a751eac8-0b434ca92139bf8-26061e51-e1000-19e24a751ebc8; _lxsdk=19e24a751eac8-0b434ca92139bf8-26061e51-e1000-19e24a751ebc8; AWPTALOS2180=; _lxsdk_cuid=19e24a75364c8-096cfaa8801d82-26061e51-e1000-19e24a75364c8; _lxsdk=19e24a75364c8-096cfaa8801d82-26061e51-e1000-19e24a75364c8; WEBDFPID=27w7u21uvy6657yv0w8373627y3xv48y80v75z246z6979586608367w-1778817806745-1778731406051EKOOAIIfd79fef3d01d5e9aadc18ccd4d0c95071243; utm_source_rg=AM%25d0YmhmF%25141%25blClvbRvTyKKGlyTfC.3l3Kbly3ATL.y.fTlGVbLKVKglgG.KKf.3KlC; eplt=J77aweWuHZ-fbRnBnRS8ipiQGa50D3A9K_21bXV7yiQmMTBfrchFXkcrpB2F9vbkDr87yMnbBBlq0_Ivs4GHDQ; eprt=i0QW7FYV4RzHvjnp2QieGMQkGARyPBnHywvfV1tOEcrSq6M2gPHqe_WWQXJahpPYbH5ucqKsuoGmMf4y2m5KHg; e_b_id_352126=531a8a5e3ed72570c6ebbfc3c8ab0dba; bizType=2; com.sankuai.meishi.fe.kdb-bsid=J77aweWuHZ-fbRnBnRS8ipiQGa50D3A9K_21bXV7yiQmMTBfrchFXkcrpB2F9vbkDr87yMnbBBlq0_Ivs4GHDQ; com.sankuai.meishi.fe.bizsettle-bsid=J77aweWuHZ-fbRnBnRS8ipiQGa50D3A9K_21bXV7yiQmMTBfrchFXkcrpB2F9vbkDr87yMnbBBlq0_Ivs4GHDQ; com.sankuai.meishi.fe.bizvisual-bsid=J77aweWuHZ-fbRnBnRS8ipiQGa50D3A9K_21bXV7yiQmMTBfrchFXkcrpB2F9vbkDr87yMnbBBlq0_Ivs4GHDQ; ecom_kdb_to_jyb_gray_flag=1; _lxsdk_s=19e24a75364-863-8e2-8dc%7C%7C9; bizType=2; edper=J77aweWuHZ-fbRnBnRS8ipiQGa50D3A9K_21bXV7yiQmMTBfrchFXkcrpB2F9vbkDr87yMnbBBlq0_Ivs4GHDQ; ecom_kdb_to_jyb_gray_flag=1; logan_session_token=sne8li54opq6egmwmdpy; com.sankuai.nibexperience.rcf.websdk_strategy=; _hc.v=25bc57a7-02a7-e2c7-9c03-b284bac00c52.1778731425; _lxsdk_s=19e24a751eb-ca9-95a-f3%7C%7C7; platformSource=1; userPlatform=PC; requestSource=dp; realAccountId=138660144; accountSource=1; accountId=76811996"

    # 产品类型ID映射
    PRD_IDS = "1,2,3,4,5,6,11,12,13,14,15,16,17,18,19,20"

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(self.HEADERS)
        self.session.headers["cookie"] = self.COOKIE

    # mtgsig（从 auth/credentials.json 读取）
    MTGSIG = '{"a1":"1.2","a2":1778731426513,"a3":"1778731404479GQQAKEEfd79fef3d01d5e9aadc18ccd4d0c95073836","a5":"XUM8dFy94iStBnNc6Bj7422ptIY3Z4Ji9jiACp6M7CwEIROs/f1J06sbObrTjpHFhpZJuXwZ5EhbmEuJ47WcxyGTyW==","a6":"hs1.6H14qYlFkNgTaC6uJjAZfO7hzM2/mQXaacK6yW0UYPkItbnNp6LS0sCQEo/hGobYaJ+LTULjl04yBaLVaj8YH66++NlIf2I9sW6cHN+X7ikRuN/QPNOuHCwc3PoZ6SK1v","a8":"7c21829712c3969fd96b3e31b231ac4b","a9":"4.2.0,7,84","a10":"21","x0":4,"d1":"fba9982040953675112e0a5bbb3cae47"}'

    def _build_mtgsig(self):
        """构建 mtgsig 参数"""
        return self.MTGSIG

    def _extract_mtgsig_from_url(self, url):
        """从URL中提取mtgsig参数"""
        match = re.search(r'mtgsig=([^&]+)', url)
        if match:
            return match.group(1)
        return None

    def get_report(self, shop_id, platform="0", begin_date=None, end_date=None, mtgsig=None):
        """
        获取交易分析报表数据

        Args:
            shop_id: 门店ID
            platform: 平台 (0=点评+美团, 1=美团, 2=点评)
            begin_date: 开始日期 (YYYY-MM-DD)
            end_date: 结束日期 (YYYY-MM-DD)
            mtgsig: mtgsig参数（可选，默认使用预设值）

        Returns:
            API原始响应数据
        """
        if mtgsig is None:
            mtgsig = self._build_mtgsig()

        # 构建URL参数
        timestamp = 1778731426513  # 使用凭证中的时间戳
        params = {
            "yodaReady": "h5",
            "csecplatform": "4",
            "csecversion": "4.2.0",
            "mtgsig": mtgsig,
        }

        # 构建请求体
        date_range = f"{begin_date},{end_date}" if begin_date and end_date else ""
        data = {
            "source": "1",
            "device": "pc",
            "date": date_range,
            "platform": platform,
            "pageType": "v5Trade",
            "optionType": "v5Trade",
            "shopIds": shop_id,
            "excludeShopIds": "",
            "cityId": "",
            "prdIds": self.PRD_IDS,
            "spuId": "",
            "pageNum": "",
            "pageSize": "",
            "sign": "",
            "fromPage": "storeKey",
            "timeStamp": str(timestamp),
        }

        response = self.session.post(
            self.BASE_URL,
            params=params,
            data=data,
            timeout=30
        )
        response.raise_for_status()
        return response.json()

    def get_metrics(self, shop_id, platform="0", begin_date=None, end_date=None):
        """
        获取交易指标（简化返回）

        Args:
            shop_id: 门店ID
            platform: 平台 (0=点评+美团, 1=美团, 2=点评)
            begin_date: 开始日期 (YYYY-MM-DD)
            end_date: 结束日期 (YYYY-MM-DD)

        Returns:
            包含各项指标的字典
        """
        response = self.get_report(shop_id, platform, begin_date, end_date)

        if response.get("code") != 0:
            raise Exception(f"API错误: {response.get('msg')}")

        # 解析body数据
        data = response.get("data", [])
        if not data:
            return {}

        body = data[0].get("body", [])

        metrics = {}
        for item in body:
            name = item.get("name", "")
            value = item.get("value", "")

            # 映射字段名
            if "下单人数" in name:
                metrics["order_person_count"] = value
            elif "下单券数" in name:
                metrics["order_ticket_count"] = value
            elif "下单金额（原价）" in name:
                metrics["order_original_amount"] = value
            elif "下单金额" in name and "原价" not in name:
                metrics["order_amount"] = value
            elif "核销人数" in name:
                metrics["redeem_person_count"] = value
            elif "核销券数" in name:
                metrics["redeem_ticket_count"] = value
            elif "核销金额（原价）" in name:
                metrics["redeem_original_amount"] = value
            elif "核销金额" in name and "原价" not in name:
                metrics["redeem_amount"] = value
            elif "退款券数" in name:
                metrics["refund_ticket_count"] = value
            elif "退款金额（原价）" in name:
                metrics["refund_original_amount"] = value

        return metrics


def test_api():
    """测试交易分析API"""
    print("=" * 60)
    print("交易分析 API 测试")
    print("=" * 60)

    client = TradeAnalysisClient()

    # 测试参数
    shop_id = "1933643130"

    print(f"\n测试参数:")
    print(f"  门店ID: {shop_id}")
    print(f"  日期范围: 2026-05-08 ~ 2026-05-14")
    print(f"  平台: 点评+美团 (2)")
    print()

    try:
        # 获取报表数据
        print("正在请求 API...")
        data = client.get_report(
            shop_id=shop_id,
            platform="0",
            begin_date="2026-05-08",
            end_date="2026-05-14"
        )

        print(f"API 返回 code: {data.get('code')}")
        print(f"API 返回 msg: {data.get('msg')}")
        print(f"API 返回 success: {data.get('success')}")
        print()

        # 打印原始响应用于调试
        print("=" * 40)
        print("原始响应数据:")
        print("=" * 40)
        print(json.dumps(data, indent=2, ensure_ascii=False))
        print()
        print()

        # 解析并显示数据
        metrics = client.get_metrics(
            shop_id=shop_id,
            platform="0",
            begin_date="2026-05-08",
            end_date="2026-05-14"
        )

        print("=" * 40)
        print("交易分析指标:")
        print("=" * 40)
        print(f"  下单人数: {metrics.get('order_person_count', 'N/A')}")
        print(f"  下单券数: {metrics.get('order_ticket_count', 'N/A')}")
        print(f"  下单金额（原价）: {metrics.get('order_original_amount', 'N/A')}")
        print(f"  下单金额: {metrics.get('order_amount', 'N/A')}")
        print(f"  核销人数: {metrics.get('redeem_person_count', 'N/A')}")
        print(f"  核销券数: {metrics.get('redeem_ticket_count', 'N/A')}")
        print(f"  核销金额（原价）: {metrics.get('redeem_original_amount', 'N/A')}")
        print(f"  核销金额: {metrics.get('redeem_amount', 'N/A')}")
        print(f"  退款券数: {metrics.get('refund_ticket_count', 'N/A')}")
        print(f"  退款金额（原价）: {metrics.get('refund_original_amount', 'N/A')}")
        print()

        print("=" * 60)
        print("API 测试成功!")
        print("=" * 60)

        return True

    except Exception as e:
        print(f"API 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_api()
    exit(0 if success else 1)
