# -*- coding: gbk -*-

from time import sleep
import speedtest   # ����speedtest_cli
import time
from openpyxl import load_workbook
#import xlsmwriter

 
#����excel��ע��·��Ҫ��ű�һ��
wb = load_workbook('test1.xlsx')
print("�Ѽ���")
#����excel��
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
        print("׼������ing..." + " " + str(t.tm_year) + "/" + str(t.tm_mon) + "/" + str(t.tm_mday) + " " + str(t.tm_hour) +":" + str(t.tm_min) + ":" + str(t.tm_sec))

        sheet[ "a"+ str(i) ] = str(t.tm_year) + "/" + str(t.tm_mon) + "/" + str(t.tm_mday) + " " + str(t.tm_hour) +":" + str(t.tm_min) + ":" + str(t.tm_sec)
        print("������д��")
        # ����ʵ������
        test = speedtest.Speedtest()
        print("�Ѵ���ʵ������")
        # ��ȡ�����ڲ��Եķ������б�
        test.get_servers()
        print("�������б��ѻ�ȡ")
        # ɸѡ����ѷ�����
        best = test.get_best_server()
 
        print("���ڲ���ing...")
 
        # �����ٶ� 
        download_speed = int(test.download() / 1024 / 1024)
        # �ϴ��ٶ�
        upload_speed = int(test.upload() / 1024 / 1024)
     
        # ������
        print("�����ٶȣ�" + str(download_speed) + " Mbits")
        print("�ϴ��ٶȣ�" + str(upload_speed) + " Mbits")
        sheet['b' + str(i)] = str(download_speed)
        sheet['c' + str(i)] = str(upload_speed)
        wb.save('test1.xlsx')
        print("д��ɹ�")
        i = i + 1
        #sleep(60)
    except:
        print("�޷����ӣ�����")
        continue