import requests
import time
from datetime import datetime, timedelta
from config import PPT_API_TOKEN,TEAM_ID  # 假设token存储在config.py中


class PPT_API():
    def __init__(self):
        self.token = PPT_API_TOKEN # 替换为实际的token
        self.teamId = TEAM_ID  # 替换为实际的teamId

    def get_users(self):
        base_url = "https://wx.pptsport.com/app/team/pageLeaderboard"
    
        # 固定参数
        params = {
            "leaderboardType": "5",
            "teamId": self.teamId,
            "pageSize": "10",
            "userName": ""
        }
        
        # 请求头
        headers = {
            "Host": "wx.pptsport.com",
            "Connection": "keep-alive",
            "content-type": "application/json",
            "Authorization": f"Bearer {self.token}",
            "SDKVersion": "3.7.12",
            "platform": "ios",
            "Accept-Encoding": "gzip,compress,br,deflate",
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_2_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 MicroMessenger/8.0.56(0x1800383b) NetType/WIFI Language/zh_CN",
            "Referer": "https://servicewechat.com/wx6989cf3c6871ac9b/579/page-frame.html"
        }
        
        all_users = []
        total_pages = None
        
        for page in range(1, 21):  # 遍历1-20页
            print(f"正在获取第 {page} 页数据...")
            
            # 更新页码参数
            params["pageIndex"] = str(page)
            
            try:
                response = requests.get(
                    base_url,
                    params=params,
                    headers=headers
                )
                
                # 检查响应状态码
                response.raise_for_status()
                
                data = response.json()

                # 检查返回码
                if data.get("code") != 200:
                    print(f"第 {page} 页请求失败，返回码: {data.get('code')}")
                    break
                # 如果是第一次请求，获取总页数
                if total_pages is None:
                    total_pages = int(data["data"]["leaderboard"]["pages"])
                    print(f"总页数: {total_pages}")
                # 如果页数大于总页数，退出循环
                if page > total_pages:
                    print(f"实际总页数只有 {total_pages} 页，将停止在总页数")
                    break
                
                # 提取用户信息
                records = data["data"]["leaderboard"]["records"]

                for user in records:
                    user_info = {
                        "userId": user["userId"],
                        # "teamId": user["teamId"],
                        "userName": user["userName"],
                        "avatar": user["avatar"],
                        "distance": user["distance"],
                        # "frequency": user["frequency"],
                        # "joinTime": user["joinTime"]
                    }
                    all_users.append(user_info)
                
                # 添加延迟，避免请求过于频繁
                time.sleep(0.5)
                
            except requests.exceptions.RequestException as e:
                print(f"第 {page} 页请求失败:", e)
                break
            except Exception as e:
                print(f"第 {page} 页处理数据时出错:", e)
                break
        
        return all_users

    def get_user_data(self, user_id, year, month):
        url = "https://api.pptsport.com/app/workout/page"
    
        # 请求参数
        params = {
            "userId": user_id,
            "teamId": self.teamId,
            "year": year,
            "month": month,
            "pageIndex": "1",
            "pageSize": "10"
        }
        
        # 请求头
        headers = {
            "Host": "wx.pptsport.com",
            "Connection": "keep-alive",
            "content-type": "application/json",
            "Authorization": f"Bearer {self.token}",
            "SDKVersion": "3.7.12",
            "platform": "ios",
            "Accept-Encoding": "gzip,compress,br,deflate",
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_2_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 MicroMessenger/8.0.56(0x1800383b) NetType/WIFI Language/zh_CN",
            "Referer": "https://servicewechat.com/wx6989cf3c6871ac9b/579/page-frame.html"
        }
        
        try:
            response = requests.get(
                url,
                params=params,
                headers=headers
            )
            
            # 检查响应状态码
            response.raise_for_status()
            
            # # 打印响应内容
            # print("请求成功!")
            # print("状态码:", response.status_code)
            # print("响应内容:", response.json())
            
            return response.json()
        
        except requests.exceptions.RequestException as e:
            print("请求失败:", e)
            return None
        
    from datetime import datetime, timedelta

    def get_weekly_data(self, user_id, target_date):
        """
        获取某周完整数据（自动处理跨月情况）
        
        Args:
            user_id: 用户ID
            target_date (str/datetime): 目标日期（如 "2025-06-20"），用于确定周范围
        Returns:
            list: 合并后的周数据（跨月则包含两个月的数据）
        """
        # 1. 计算周的起始日期（周一）和结束日期（周日）
        if isinstance(target_date, str):
            target_date = datetime.strptime(target_date, "%Y-%m-%d")
        monday = target_date - timedelta(days=target_date.weekday())
        sunday = monday + timedelta(days=6)
        
        # 2. 检查是否跨月
        if monday.month != sunday.month:
            # 跨月：请求两个月份的数据
            month1 = monday.strftime("%Y%m")
            month2 = sunday.strftime("%Y%m")
            data1 = self.get_user_data(user_id, month1[:4], month1[4:])
            data2 = self.get_user_data(user_id, month2[:4], month2[4:])
            # 合并 records 数据
            combined_records = (
                data1.get("data", {}).get("records", []) + 
                data2.get("data", {}).get("records", [])
            )
            return {
                "data": {
                    "records": combined_records
                }
            }
        else:
            # 未跨月：请求单月数据
            month = monday.strftime("%Y%m")
            return self.get_user_data(user_id, month[:4], month[4:])