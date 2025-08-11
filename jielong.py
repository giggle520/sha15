import requests
import time
from collections import defaultdict
from config import ACT_ID,JIELONG_API_TOKEN


class JIELONG_API():
    def __init__(self):
        self.actId = ACT_ID
        self.JIELONG_API_TOKEN = JIELONG_API_TOKEN
    
    def get_jielong_info(self):
        url = "https://apipro.qunjielong.com/order-cert/api/mina/activity-order/get_group_act_last_orders"
        
        headers = {
            "Host": "apipro.qunjielong.com",
            "Connection": "keep-alive",
            "content-type": "application/json",
            "appid": "wx059cd327295ab444",
            "sceneCode": "1089",
            "uid": "63510141",
            "version": "3.0.0",
            "client-version": "5.9.85",
            "device-type": "5",
            "actId": self.actId,
            "scene-flow": "seq-detail,15f0c67ab374caeca66a57e663a076a5",
            "Authorization": self.JIELONG_API_TOKEN,
            "mini-route": "pro/pages/seq-detail/detail-sign-up/detail-sign-up",
            "companyId": "190",
            "feature-tag": "f0000",
            "Accept-Encoding": "gzip,compress,br,deflate",
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_2_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 MicroMessenger/8.0.56(0x1800383b) NetType/4G Language/zh_CN",
            "Referer": "https://servicewechat.com/wx059cd327295ab444/2104/page-frame.html"
        }
        
        # 初始化分组字典
        group_dict = defaultdict(list)
        
        # 获取所有页数据
        total_pages = None
        current_page = 1
        
        while True:
            # 准备请求数据
            data = {
                "actId": self.actId,
                "pageParam": {
                    "page": current_page,
                    "pageSize": 10
                },
                "sortMethod": 10,
                "lookAll": False
            }
            
            try:
                print(f"正在获取第 {current_page} 页数据...")
                response = requests.post(
                    url,
                    headers=headers,
                    json=data
                )
                
                # 检查响应状态码
                response.raise_for_status()
                
                # 解析JSON数据
                result = response.json()
                
                # 检查返回码
                if result.get("code") != 200:
                    print(f"第 {current_page} 页请求失败，返回码: {result.get('code')}")
                    break
                
                # 如果是第一次请求，获取总页数
                if total_pages is None:
                    total_pages = result["data"]["pageInfo"]["allPage"]
                    print(f"总页数: {total_pages}")
                
                # 解析数据
                for item in result["data"]["entityList"]:
                    try:
                        # 获取组别信息
                        apply_items = item["applyActivityFeedInfo"]["applyItemList"]
                        group_name = apply_items[0]["applyItemName"] if apply_items else "未知组别"
                        
                        # 获取人名信息
                        additional_items = item["applyActivityFeedInfo"]["additionalItemList"]
                        name = "未知姓名"
                        for attr in additional_items:
                            # if attr["attrText"] == "姓名":
                            if "姓名" in attr["attrText"]:
                                name = attr["itemValue"]["mainInfo"].strip()
                                break
                        
                        # 添加到分组字典
                        group_dict[group_name].append(name)
                    except KeyError as e:
                        print(f"第 {current_page} 页数据解析错误: 缺少键 {e}")
                        continue
                
                # 检查是否还有下一页
                if current_page >= total_pages:
                    break
                    
                current_page += 1
                # 添加延迟，避免请求过于频繁
                time.sleep(0.5)
                
            except requests.exceptions.RequestException as e:
                print(f"第 {current_page} 页请求失败:", e)
                break
            except Exception as e:
                print(f"第 {current_page} 页处理数据时出错:", e)
                break
        
        # print(group_dict)
        # 输出结果
        # print("\n报名情况统计:")
        # for group, names in group_dict.items():
        #     print(f"\n【{group}】组 ({len(names)}人):")
        #     for i, name in enumerate(names, 1):
        #         print(f"{i}. {name}")

        # 按照公里数从大到小排序
        sorted_groups = sorted(group_dict.items(), key=lambda x: int(x[0].replace('公里', '')), reverse=True)

        # 转换为有序字典（可选）
        from collections import OrderedDict
        ordered_group_dict = OrderedDict(sorted_groups)
        
        return ordered_group_dict

if __name__ == "__main__":
    JIELONG_API().get_jielong_info()
