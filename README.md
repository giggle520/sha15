## 该项目主要实现了以下功能：
**1.自动统计体能训练接龙情况**  
**2.自动统计跑跑堂跑团成员打卡情况**  
**3.将统计结果导出为目标Excel文件** 

### 文件功能介绍：  
**main.py**: 主程序    
**config.py**: 配置文件，每次使用前需要事先抓取请求包里的token等,苹果使用Stream,安卓使用HttpCanary  
**jielong.py**: 获取体能训练接龙情况的脚本    
**ppt.py**: 获取跑跑堂跑团成员打卡情况的脚本      
**write_to_excel.py**： 实现写入生成excel的脚本    

### 使用步骤： 
1.首次使用先安装依赖: pip install -r requirements.txt  
2.每次使用前更新配置文件config.py中的配置信息    
3.运行方法：python main.py    



