# -*- coding: gbk -*-

from time import sleep
import speedtest   # 导入speedtest_cli
import time
from openpyxl import load_workbook
#import xlsmwriter

 
#加载excel，注意路径要与脚本一致
wb = load_workbook('test1.xlsx')
print("已激活")
#激活excel表
sheet = wb.active
sheet['a1'] = 'zeit'
sheet['b1'] = 'unterladen'
sheet['c1'] = 'aufladen'
t= time.localtime()
i=257

while i > 1 :
    try:
        #print(t)  t.tm_year, t.tm_mon, t.tm_mday,t.tm_hour, t.tm_min, t.tm_sec,1900/1/10 0:00:00
        t= time.localtime()
        print("准备测试ing..." + " " + str(t.tm_year) + "/" + str(t.tm_mon) + "/" + str(t.tm_mday) + " " + str(t.tm_hour) +":" + str(t.tm_min) + ":" + str(t.tm_sec))

        sheet[ "a"+ str(i) ] = str(t.tm_year) + "/" + str(t.tm_mon) + "/" + str(t.tm_mday) + " " + str(t.tm_hour) +":" + str(t.tm_min) + ":" + str(t.tm_sec)
        print("日期已写入")
        # 创建实例对象
        test = speedtest.Speedtest()
        print("已创建实例对象")
        # 获取可用于测试的服务器列表
        test.get_servers()
        print("服务器列表已获取")
        # 筛选出最佳服务器
        best = test.get_best_server()
 
        print("正在测试ing...")
 
        # 下载速度 
        download_speed = int(test.download() / 1024 / 1024)
        # 上传速度
        upload_speed = int(test.upload() / 1024 / 1024)
     
        # 输出结果
        print("下载速度：" + str(download_speed) + " Mbits")
        print("上传速度：" + str(upload_speed) + " Mbits")
        sheet['b' + str(i)] = str(download_speed)
        sheet['c' + str(i)] = str(upload_speed)
        wb.save('test1.xlsx')
        print("写入成功")
        i = i + 1
        #sleep(60)
    except:
        print("无法连接，重试")
        continue