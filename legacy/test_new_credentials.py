"""
测试新获取的 Cookie 和 mtgsig 是否有效
"""
import requests

url = "https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport"

# 从 credentials.json 读取的新凭证
mtgsig = "%7B%22a1%22%3A%221.2%22%2C%22a2%22%3A1778731171142%2C%22a3%22%3A%221778731138289KKYWEYQfd79fef3d01d5e9aadc18ccd4d0c95071644%22%2C%22a5%22%3A%22at4eUhAkuuCDr6Vfw%2FJF%2Fg0U4TZhd6k%2Ff8MvdenXO5BYxvgySmWFFFbaqAHLGmj99haDeYXVMBKFFGnb3ynNl3Zfgc%3D%3D%22%2C%22a6%22%3A%22hs1.6F5jb%2Fm9VJvZJfUZaM%2FHe6AMWhKOLvnnWq4OUXVOVuLrjXXH%2B%2FjpwB%2FhTAHWsJZ9RbybgHBN6OXhohQ4vATRp7VPA%2BkIcAeHE4EnuPv4xS2nZcGtRyyVwB6xoH6%2B1ePDR%22%2C%22a8%22%3A%22dd416d94ac8af39a5af4b87e04974a6e%22%2C%22a9%22%3A%224.2.0%2C7%2C137%22%2C%22a10%22%3A%2275%22%2C%22x0%22%3A4%2C%22d1%22%3A%223a9412b2ff706e73bcb78a0fc5e47cd3%22%7D"

cookie = "AWPTALOS32837=; WEBDFPID=1778731138289KKYWEYQfd79fef3d01d5e9aadc18ccd4d0c95071644-1778731138289-1778731138289KKYWEYQfd79fef3d01d5e9aadc18ccd4d0c95071644; _lxsdk_cuid=19e24a34234c8-08b646726f7e87-26061e51-e1000-19e24a34234c8; _lxsdk=19e24a34234c8-08b646726f7e87-26061e51-e1000-19e24a34234c8; AWPTALOS2180=; _lxsdk_cuid=19e24a34376c8-0af784061a11b98-26061e51-e1000-19e24a3437610; _lxsdk=19e24a34376c8-0af784061a11b98-26061e51-e1000-19e24a3437610; WEBDFPID=32x9y11299995541z93854557722177680v75z656y7979584008vzw1-1778817540505-1778731139867CGWUOCIfd79fef3d01d5e9aadc18ccd4d0c95071847; utm_source_rg=AM%25f7rq7qX%25433%25DlX5HPPl5555OOgPA5DcOgOO11llP11McvW1OAMOMH1515OcgvvcWA4P; eplt=1clO-PkUpBGlz7HLT57EYWeZoGfAyB6knEC5Z9GKz5w8mdELG-6lgB9BAzoeYbxPcRQCamzcG-2jzkjrk1thtA; eprt=6tlSfgWNo_NXQNUNx1gyuzcpu_0001oly6P2AIJ1C0kV_yN9hz7dKxIcA_BnI_a5ctH83KN8PL0Ptx3QyZJ8Nw; e_b_id_352126=531a8a5e3ed72570c6ebbfc3c8ab0dba; ecom_kdb_to_jyb_gray_flag=1; bizType=2; com.sankuai.meishi.fe.kdb-bsid=1clO-PkUpBGlz7HLT57EYWeZoGfAyB6knEC5Z9GKz5w8mdELG-6lgB9BAzoeYbxPcRQCamzcG-2jzkjrk1thtA; com.sankuai.meishi.fe.bizsettle-bsid=1clO-PkUpBGlz7HLT57EYWeZoGfAyB6knEC5Z9GKz5w8mdELG-6lgB9BAzoeYbxPcRQCamzcG-2jzkjrk1thtA; com.sankuai.meishi.fe.bizvisual-bsid=1clO-PkUpBGlz7HLT57EYWeZoGfAyB6knEC5Z9GKz5w8mdELG-6lgB9BAzoeYbxPcRQCamzcG-2jzkjrk1thtA; _lxsdk_s=19e24a34377-a57-928-dca%7C%7C9; bizType=2; edper=1clO-PkUpBGlz7HLT57EYWeZoGfAyB6knEC5Z9GKz5w8mdELG-6lgB9BAzoeYbxPcRQCamzcG-2jzkjrk1thtA; ecom_kdb_to_jyb_gray_flag=1; fromEntry=1; logan_session_token=pn3rjehou61im6sefw1e; com.sankuai.nibexperience.rcf.websdk_strategy=; _hc.v=c41a92aa-e6bf-8d29-64cc-9ae48df5021b.1778731171; _lxsdk_s=19e24a34235-fe7-1ad-7bd%7C%7C7; platformSource=1; userPlatform=PC; requestSource=dp; realAccountId=138660144; accountSource=1; accountId=76811996"

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
    "mtgsig": mtgsig
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://e.dianping.com/shopdiy-node/report",
    "Cookie": cookie
}

print("Testing new credentials from credentials.json...")
response = requests.get(url, params=params, headers=headers)

print(f"Status: {response.status_code}")
print(f"Response (first 500 chars):\n{response.text[:500]}")