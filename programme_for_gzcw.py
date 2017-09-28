#-*- coding:utf-8 -*-

"""
广州公共资源交易网
"""

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
import xlwt
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import random

# def get_content(url):
# 	driver = webdriver.PhantomJS()
# 	driver.set_window_size(800,600)

# 	driver.get(url)
# 	driver.find_element_by_link_text(u"政府采购").click()
# 	# driver.find_element_by_css_selector("span > dd").click()
# 	driver.find_element_by_xpath("/x:html/x:body/x:div[2]/x:div[3]/x:div[1]/x:div[2]/x:div[1]/x:div[5]/x:a").click()
# 	driver.find_element_by_id("searchvalue").clear()
# 	driver.find_element_by_id("searchvalue").send_keys(u"医院")
# 	driver.find_element_by_id("img1").click()

# 	return driver
class Programme_gzcw():
	programme_name_list = []
	dateindex = []

	def get_content2_gzcw(self,url,keyword):
		"""
		打开网址
		搜索关键词
		"""
		driver = webdriver.PhantomJS()
		driver.set_window_size(800,600)

		driver.get(url)
		driver.find_element_by_id("searchvalue").clear()
		driver.find_element_by_id("searchvalue").send_keys(keyword)
		driver.find_element_by_id("img1").click()

		return driver


	def get_web_gzcw(self,info,starttime,endtime,state):
		"""
		获取项目网址
		实现翻页
		"""
		
		corrt_weblist = []
		if state is True:
			while True:
				weblist=[]
				programmelist = []
				soup = BeautifulSoup(info.page_source)

				website = soup.find_all(href=re.compile("layout3"))
				dateall = soup.find_all('td',text = re.compile("\d{4}-\d+-\d+"))

				for web in website:
					weblist.append(web['href'])
					programmelist.append(web.get_text())
				weblist.remove(weblist[0])
				weblist.remove(weblist[0])
				programmelist.remove(programmelist[0])
				programmelist.remove(programmelist[0])

				for web_unit in weblist:
					corrt_weblist.append(web_unit)
				for programmelist_unit in programmelist:
					self.programme_name_list.append(programmelist_unit)

				for date in dateall:
					datestr = date.get_text()
					dateint = int(datestr[0]+datestr[1]+datestr[2]+datestr[3]+datestr[5]+datestr[6]+datestr[8]+datestr[9])
					self.dateindex.append(dateint)

				pageall = soup.find('span',text = re.compile(u"共\d*页")).get_text()
				pagenow = soup.find('span',text = re.compile(u"第\d*页")).get_text()
				pageallint = re.findall("\d+",pageall)
				pagenowint = re.findall("\d+",pagenow)
				if int(pagenowint[0]) < int(pageallint[0]):
					info.find_element_by_link_text("下一页").click()
					print("翻1页")
					time.sleep(random.randint(3,5))
				else:
					break

		else:
			while True:
				datelist = []
				weblist = []
				site = 0
				programmelist = []

				soup = BeautifulSoup(info.page_source)

				dateall = soup.find_all('td',text = re.compile("\d{4}-\d+-\d+"))
				website = soup.find_all(href=re.compile("layout3"))

				for web in website:
					weblist.append(web['href'])
					programmelist.append(web.get_text())
				weblist.remove(weblist[0])
				weblist.remove(weblist[0])
				programmelist.remove(programmelist[0])
				programmelist.remove(programmelist[0])

				for date in dateall:
					datestr = date.get_text()
					dateint = int(datestr[0]+datestr[1]+datestr[2]+datestr[3]+datestr[5]+datestr[6]+datestr[8]+datestr[9])
					datelist.append(dateint)

				for datelist_unit in datelist:
					if starttime<= datelist_unit <= endtime:
						self.dateindex.append(datelist_unit)
						self.programme_name_list.append(programmelist[site])
						corrt_weblist.append(weblist[site])
					else:
						None
					site +=1

				if starttime > datelist[len(datelist)-1]:
					break
				else:
					pageall = soup.find('span',text = re.compile(u"共\d*页")).get_text()
					pagenow = soup.find('span',text = re.compile(u"第\d*页")).get_text()
					pageallint = re.findall("\d+",pageall)
					pagenowint = re.findall("\d+",pagenow)
					if int(pagenowint[0]) < int(pageallint[0]):
						info.find_element_by_link_text("下一页").click()
						print("翻1页")
						time.sleep(random.randint(3,5))
					else:
						break

		return corrt_weblist

	def get_agentcompany_gzcw(self,list_table):
		return "广州公共资源交易网"
	# 	i = 0
	# 	site = 0
	# 	for message in list_table:
	# 		try:
	# 			message.index(u"采购代理机构")
	# 			site = i
	# 		except:
	# 			i +=1
	# 	# if site !=0:
	# 	# 	return list_table[site]
	# 	# else:
	# 	# 	return None
	# 	return i


	def get_buyer_gzcw(self,list_table):
		"""
		获取采购人
		"""
		i = 0
		site = 0
		for message in list_table:
			try:
				message.index(u"采购人名称")
				site = i
			except:
				i +=1
		if site !=0:
			buyer = list_table[site]
			return buyer[6:]
		else:
			return None

	def get_showtime_gzcw(self,list_table):
		"""
		获取开标时间
		"""
		i = 0
		site = 0
		j = 0
		showtime = None
		for message in list_table:
			try:
				message.index(u"和开标时间")
				site = i
			except:
				i +=1

		if site !=0:
			if len(list_table[site+1]) < 2:
				showtime =  list_table[site +2]
			else:
				showtime = list_table[site+1]
		else:
			for message2 in list_table:
				try:
					message2.index(u"及开标时间")
					site = j
				except:
					j +=1
			if site !=0:
				if len(list_table[site+1]) <2:
					showtime =  list_table[site+2]
				else:
					showtime =  list_table[site+1]
			else:
				None
		try:
			showtimefind = re.findall(r"\d+\年\d+\月\d+\日",showtime)
		except:
			showtimefind = ""

		if len(showtimefind) !=0:
			return showtimefind[0]
		else:
			return None

	def get_account_gzcw(self,list_table):
		"""
		获取项目预算
		"""
		i = 0
		site = 0
		for message in list_table:
			try:
				message.index(u"预算")
				site = i
			except:
				i +=1

		if site !=0:
			if len(list_table[site+1]) <2:
				return list_table[site+2]
			else:
				return list_table[site+1]
		else:
			return None

	def get_money_gzcw(self,list_table):
		i = 0
		site = 0
		for mes in list_table:
			try:
				mes.index(u"中标金额")
				site = i
			except:
				try:
					mes.index(u"成交金额")
					site = i
				except:
					try:
						mes.index(u"总报价")
						site = i
					except:
						i +=1


		if site !=0:
			moneylist = re.findall(r"\d+\.?\d*",list_table[site])
			if len(moneylist) !=0:
				money = float(moneylist[0])
				return money
			else:
				return None

		else:
			return "无法获取"

	def get_programme_destinate_gzcw(self,list_table):
		"""
		获取采购内容
		"""
		i = 0
		site = 0
		for message in list_table:
			try:
				message.index(u"采购内容：")
				site = i
			except:
				i +=1

		if site !=0:
			return list_table[site+1]
		else:
			return None

	def get_detail_gzcw(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
		"""
		总调用方法
		写入EXCEL
		"""
		
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
			html = urllib.request.urlopen('http://www.gzggzy.cn' + web)
			content = html.read()
			html.close()
			"""打开项目网址"""
			message = []

			soup = BeautifulSoup(content)

			message_all = soup.find_all('div', class_="xx-text")
			for message_part in message_all:
				for message_unit in message_part.stripped_strings:
					message.append(message_unit)

			# print(get_title(soup),get_beginningtime(message),get_buyer(message),get_showtime(message),get_account(message),get_programme_destinate(message))

			if state1 is True:
				sheet.write(excel_count,5,self.programme_name_list[excel_count-1])
			else:
				None
			if state2 is True:	
				sheet.write(excel_count,12,self.get_account_gzcw(message))
			else:
				None
			if state3 is True:
				sheet.write(excel_count,0,self.dateindex[excel_count-1])
			else:
				None
			if state4 is True:	
				sheet.write(excel_count,3,self.get_agentcompany_gzcw(message))
			else:
				None
			if state5 is True:
				sheet.write(excel_count,1,"http://www.gdgpo.gov.cn" + web)
			else:
				None
			if state6 is True:
				sheet.write(excel_count,11,self.get_money_gzcw(message))
			else:
				None
			if state7 is True:
				sheet.write(excel_count,7,self.get_showtime_gzcw(message))
			else:
				None
			if state8 is True:
				sheet.write(excel_count,4,self.get_buyer_gzcw(message))
			else:
				None

			excel_count +=1

		wbk.save(filename)

