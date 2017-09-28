#-*- coding:utf-8 -*-

"""
"""

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
import xlwt
from selenium import webdriver
import datetime
import time
import random

# .decode('utf-8')

class Programme_jiamu():

	def get_html(self,starttime,endtime,filename):

		"""
		打开Excel
		"""
		wbk = xlwt.Workbook()
		sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)

		excel_count = 0
		sheet.write(0,0,"品种")
		sheet.write(0,1,"单价")
		sheet.write(0,2,"日期")

		begin = datetime.date(starttime[0],starttime[1],starttime[2])
		end = datetime.date(endtime[0],endtime[1],endtime[2])
		
		d = begin
		delta = datetime.timedelta(days=1)
		date_count = 1
		while d <= end:

			date = d.strftime("%Y-%m-%d")

			storename = []
			storeaccount = []
			store_list = []
			group_count = 0
			excel_count = 0

			# html = urllib.request.urlopen('http://121.8.226.252/basic/sendReportInfoes/sendReportDeatil?srhSrId=2&srhReportId=15&srhCycle=' + str(date),timeout = 30)
			# time.sleep(random.randint(3,6))
			# content = html.read()
			# html.close()
			content = webdriver.PhantomJS()
			content.set_page_load_timeout(30)
			content.set_script_timeout(30)
			content.set_window_size(800,600)
			try:
				content.get('http://121.8.226.252/basic/sendReportInfoes/sendReportDeatil?srhSrId=2&srhReportId=15&srhCycle=' + str(date))
			except:
				content.get('http://121.8.226.252/basic/sendReportInfoes/sendReportDeatil?srhSrId=2&srhReportId=15&srhCycle=' + str(date))

			time.sleep(random.randint(3,6))

			soup = BeautifulSoup(content.page_source)


			table = soup.find_all('td', class_ = "null")

			for store in table:
				store_list.append(store.get_text())

			while group_count < len(store_list)/3:
				if store_list[group_count*3] != "":
					storename.append(store_list[group_count*3])

				if store_list[(group_count+1)*3-1] != "":
					storeaccount.append(float(store_list[(group_count+1)*3-1]))
				group_count +=1

			while excel_count < len(storename):

				sheet.write(date_count,0,storename[excel_count])
				sheet.write(date_count,1,storeaccount[excel_count])
				sheet.write(date_count,2,str(date))
				excel_count +=1
				date_count +=1

			d += delta
			print(date)
			content.quit()  
			time.sleep(1)

		wbk.save(filename)
