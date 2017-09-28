#-*- coding:utf-8 -*-

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
from selenium import webdriver
import xlwt
import time
import random

class Programme_gdhl():
	datelist = []
	programme_namelist = []

	def get_content_gdhl(self,url,typetext,keyword):
		gdhl = webdriver.PhantomJS()
		gdhl.set_window_size(800,600)
		gdhl.get(url)

		gdhl.find_element_by_css_selector("body").click()
		gdhl.find_element_by_xpath("//map[@id='Map']/area[5]").click()
		gdhl.find_element_by_link_text("招标公告").click()

		gdhl.find_element_by_id("MXMMC").click()
		gdhl.find_element_by_id("MXMMC").clear()
		gdhl.find_element_by_id("MXMMC").send_keys(keyword)
		if typetext == "招标类":
			if not gdhl.find_element_by_xpath("//select[@id='MGGLB']//option[2]").is_selected():
				gdhl.find_element_by_xpath("//select[@id='MGGLB']//option[2]").click()
		else:
			if not gdhl.find_element_by_xpath("//select[@id='MGGLB']//option[3]").is_selected():
				gdhl.find_element_by_xpath("//select[@id='MGGLB']//option[3]").click()
		
		gdhl.find_element_by_id("button").click()
		time.sleep(2)

		return gdhl

	def get_web_gdhl(self,info,starttime,endtime,state):
		weblist = []
		pagecount = 0

		if state is True:
			while True:
				meslist = []
				webindex = []
				site = 0

				soup = BeautifulSoup(info.page_source)
				weball = soup.find_all('a',href = re.compile("gonggao-mx"))
				classlist = soup.find_all('span',class_ = "STYLE43")
				pagedown_index = soup.find('td', text = u"尾页")
				pagedown_index_a = pagedown_index.find('a')
				pagedown_judge = re.findall(r"pageno=\d+",pagedown_index_a['href'])
				pagecount_max = int(pagedown_judge[0][7:])

				for web in weball:
					weblist.append(web['href'])

				for classlist_unit in classlist:
					meslist.append(classlist_unit.get_text())

				while site < len(meslist)/3:
					self.datelist.append(meslist[site*3 +2])
					self.programme_namelist.append(meslist[site*3 +1])
					site +=1

				if pagecount == pagecount_max-1:
					info.quit()
					print("关闭浏览器")
					break
				else:
					info.find_element_by_link_text("下一页").click()
					print("翻一页")
					time.sleep(random.randint(3,5))
					pagecount +=1

		else:
			while True:
				meslist = []
				webindex = []
				dateindex = []
				programme_nameindex = []
				site = 0
				i = 0

				soup = BeautifulSoup(info.page_source)
				weball = soup.find_all('a',href = re.compile("gonggao-mx"))
				classlist = soup.find_all('span',class_ = "STYLE43")
				pagedown_index = soup.find('td', text = u"尾页")
				pagedown_index_a = pagedown_index.find('a')
				pagedown_judge = re.findall(r"pageno=\d+",pagedown_index_a['href'])
				pagecount_max = int(pagedown_judge[0][7:])

				for web in weball:
					webindex.append(web['href'])

				for classlist_unit in classlist:
					meslist.append(classlist_unit.get_text())

				while site < len(meslist)/3:
					dateindex.append(meslist[site*3 +2])
					programme_nameindex.append(meslist[site*3 +1])
					site +=1

				for date in dateindex:
					dateint = int(date[0]+date[1]+date[2]+date[3]+date[5]+date[6]+date[8]+date[9])
					if starttime <= dateint <= endtime:
						self.datelist.append(date)
						self.programme_namelist.append(programme_nameindex[i])
						weblist.append(webindex[i])
					else:
						None
					i +=1

				if pagecount == pagecount_max-1:
					info.quit()
					print("关闭浏览器")
					break
				else:
					datelast = dateindex[len(dateindex)-1]
					datelastint = int(datelast[0]+datelast[1]+datelast[2]+datelast[3]+datelast[5]+datelast[6]+datelast[8]+datelast[9])
					if starttime < datelastint:
						info.find_element_by_link_text("下一页").click()
						print("翻一页")
						time.sleep(random.randint(3,5))
						pagecount +=1
					else:
						info.quit()
						print("关闭浏览器")
						break
						
		return weblist

	def get_account_gdhl(self,messagelist):
		site = 0
		i = 0
		account = None

		for message in messagelist:
			try:
				message.index(u"采购项目预算金额")
				site = i
			except:
				try:
					message.index(u"最高限价")
					site = i
				except:
					i +=1

		if site != 0:
			return messagelist[site]
		else:
			return None

	def get_showtime_gdhl(self,messagelist):
		site = 0
		i = 0
		showtime = None

		for message in messagelist:
			try:
				message.index(u"开标时间")
				site = i
			except:
				i +=1
		if site != 0:
			showtime = re.findall("\d+年\d+月\d+日",messagelist[site])
			if len(showtime) !=0:
				return showtime[0]
			else:
				return None
		else:
			return "未找到"

	def get_buyer_gdhl(self,messagelist):
		site = 0
		i = 0
		buyer = None

		for message in messagelist:
			try:
				message.index(u"采购人：")
				site = i
			except:
				try:
					message.index(u"采购单位：")
					site = i
				except:
					i +=1
		if site !=0:
			return messagelist[site]
		else:
			return "未找到"

	def get_agent_gdhl(self,messagelist):
		return "广东华伦招标有限公司"

	def get_money_gdhl(self,soup):
		mlist = []
		outcome = None
		try:
			money = soup.find('table',class_=re.compile("MsoNormalTable")).stripped_strings
			for money_unit in money:
				mlist.append(money_unit)
			outcome = mlist[4]
		except:
			None
		if outcome is not None:
			if len(outcome) != 0:
				return outcome
			else:
				return "未找到"
		else:
			return "未找到"

	# def test(self,WBall):
	# 	for web in WBall:
	# 		html = urllib.request.urlopen("http://www.gdhualun.com.cn/"+web,timeout = 30)
	# 		content = html.read()
	# 		html.close()

	# 		message = []

	# 		soup = BeautifulSoup(content)

	# 		print(self.get_money_gdhl(soup))


	def get_detail_gdhl(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
		wbk = xlwt.Workbook()
		sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)

		sheet.write(0,0,"招标公告日期")
		sheet.write(0,1,"链接")
		sheet.write(0,2,"地区")
		sheet.write(0,3,"招标机构")
		sheet.write(0,4,"采购单位")
		sheet.write(0,5,"项目名称")
		sheet.write(0,6,"采购内容")
		sheet.write(0,7,"开标日期")
		sheet.write(0,8,"中标公告时间")
		sheet.write(0,9,"中标公司")
		sheet.write(0,10,"总金额")
		sheet.write(0,11,"中标金额")
		sheet.write(0,12,"预算（元）")
		sheet.write(0,13,"中标公告链接")

		excel_count = 1

		for web in WBall:
			html = urllib.request.urlopen("http://www.gdhualun.com.cn/"+web,timeout = 30)
			content = html.read()
			html.close()

			message = []

			soup = BeautifulSoup(content)

			message_all = soup.find_all('p',class_ = "MsoNormal")
			for message_part in message_all:
				message.append(message_part.get_text())

			if state1 is True:
				sheet.write(excel_count,5,self.programme_namelist[excel_count-1])
			else:
				None
			if state2 is True:
				sheet.write(excel_count,12,self.get_account_gdhl(message))
			else:
				None
			if state3 is True:
				sheet.write(excel_count,0,self.datelist[excel_count-1])
			else:
				None
			if state4 is True:
				sheet.write(excel_count,3,self.get_agent_gdhl(message))
			else:
				None
			if state5 is True:
				sheet.write(excel_count,1,"http://www.gdhualun.com.cn/"+web)
			else:
				None
			if state6 is True:
				sheet.write(excel_count,11,self.get_money_gdhl(soup))
			else:
				None
			if state7 is True:
				sheet.write(excel_count,7,self.get_showtime_gdhl(message))
			else:
				None
			if state8 is True:
				sheet.write(excel_count,4,self.get_buyer_gdhl(message))
			else:
				None

			excel_count +=1

		wbk.save(filename)

