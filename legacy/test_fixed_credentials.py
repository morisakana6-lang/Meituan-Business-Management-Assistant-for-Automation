"""
使用用户之前有效的固定 Cookie 和 mtgsig 测试 API 请求
"""
import requests

url = "https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport"

params = {
    "dimension": "shop",
    "beginDate": "2026-05-07",
    "endDate": "2026-05-13",
    "platform": "0",
    "compareEnabled": "0",
    "compareBeginDate": "2026-04-30",
    "compareEndDate": "2026-05-06",
    "objectUnit": "account",
    "groupUnit": "",
    "timeUnit": "day",
    "shopIds": "0",
    "launchIds": "0",
    "launchPremiumIds": "0",
    "planIds": "0",
    "tabIds": "T30001,T30047,T30002,T30003,T30020,T30049,T30005,T30006,T30007,T30039,T30009,T30083",
    "reportFunctionType": "2",
    "yodaReady": "h5",
    "csecplatform": "4",
    "csecversion": "4.2.0",
    "mtgsig": "%7B%22a1%22%3A%221.2%22%2C%22a2%22%3A1778726430419%2C%22a3%22%3A%22x19u036z5yvx5966y9uzyx673y7yy46w80v9yyx3039979589856295u%22%2C%22a5%22%3A%22CyLj%2Bs8QH2FMspSNvIjkbzUMeki%2B%2BUjt28sMWY9yWkFS1%2BTFouyBkskoEukAJ35pf%2BZLSUQAF2N%2F5mWbfxScxceQXc%3D%3D%22%2C%22a6%22%3A%22hs1.6jUvGqAYjCn67Hfk8Ocb9%2Btm24E6h5Rce43AtEQgJ52TlTXlnybA5Xiz0XszzR5v6xJqQbtKhqDIt7Wt%2BRz%2F%2Fm%2BDcdKA2Ci4tou2k8MJLpNastJ%2B52dy%2FgeVZmWi5VHkj%22%2C%22a8%22%3A%22b142050f7850faa3dae6fc5d4be0977f%22%2C%22a9%22%3A%224.2.0%2C7%2C206%22%2C%22a10%22%3A%220a%22%2C%22x0%22%3A4%2C%22d1%22%3A%22e1b03cf2acbe698ce3b343594d0e555c%22%7D"
}

# 使用用户在 test_meituan_request.py 中的固定 Cookie
cookies = "bizType=2; edper=m0xl6tJlBrLZtawvA2TuDuKCqfQZ5V1J8QbnUYKe4434k_bYrDvGzbsd8E67_TF874_JU6P2GOCOgR7HQTex6w; ecom_kdb_to_jyb_gray_flag=1; _gw_ab_call_15533_110=TRUE; _gw_ab_15533_110=348; _lxsdk_cuid=19e0bbc2b74c8-0699251930556e8-26061151-1fa400-19e0bbc2b74c8; _lxsdk=19e0bbc2b74c8-0699251930556e8-26061151-1fa400-19e0bbc2b74c8; _hc.v=312c8d48-84cd-6a03-1dcd-f729deac14b8.1778313343; merchantCategoryID=159; _gw_ab_call_30962_44=TRUE; _gw_ab_30962_44=197; AWPTALOS30702=; _gw_ab_call_49238_17=TRUE; _gw_ab_49238_17=731; AWPTALOS32837=; WEBDFPID=x19u036z5yvx5966y9uzyx673y7yy46w80v9yyx3039979589856295u-1778553616989-1778313341826CIAMSKCfd79fef3d01d5e9aadc18ccd4d0c95072334; utm_source_rg=AM%25f0gdldv%25329%25SvDJU2rQB0iSBDrr0DJQ0SrE20E003rF4UiD00S2U2DDEDB4D4BrzDBJ; com.sankuai.adplaunchreport.fe.isomorph_strategy=; shopOperationSource=0; com.sankuai.adplaunchcenter.menu.pc_strategy=; com.sankuai.adplaunchcenter.cpc.pc_strategy=; com.sankuai.adplaunchcenter.fullsite.isomorph_strategy=; com.sankuai.adplaunchcenter.promo.isomorph_strategy=; PROMO_THEME_VALUE=mt-yellow; com.sankuai.adplaunchcenter.cpm.pc_strategy=; JSESSIONID=4964A7EAF7E6B8754B14673CA14DF97B; mpmerchant_portal_shopid=1933643130; merchantBookShopID=1933643130; fromEntry=1; logan_session_token=66jgaqvl5ibdrbkgeiv6; _lxsdk_s=19e245a9c85-742-439-2b9%7C%7C76; platformSource=1; userPlatform=PC; requestSource=dp; realAccountId=138660144; accountSource=1; accountId=76811996"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    "Referer": "https://e.dianping.com/shopdiy-node/report",
    "Cookie": cookies
}

print("正在测试用户之前的固定 Cookie 和 mtgsig...")
response = requests.get(url, params=params, headers=headers)

print(f"\n状态码: {response.status_code}")
print(f"响应内容:\n{response.text[:2000]}")