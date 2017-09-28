#-*- coding:utf-8 -*-

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
from selenium import webdriver
import xlwt
import time
import random

class Programme_gdcl():
	datelist = []

	def get_content_gdcl(self,url,typetext,keyword):
		gdcl = webdriver.PhantomJS()
		gdcl.set_window_size(800,600)
		gdcl.get(url)

		# gdcl.find_element_by_css_selector("body.bodybg").click()
		# gdcl.find_element_by_link_text(u"信息公告").click()
		if typetext == "招标公告":
			if not gdcl.find_element_by_xpath("//select[@id='catid_1']//option[2]").is_selected():
				gdcl.find_element_by_xpath("//select[@id='catid_1']//option[2]").click()
		else:
			if not gdcl.find_element_by_xpath("//select[@id='catid_1']//option[4]").is_selected():
				gdcl.find_element_by_xpath("//select[@id='catid_1']//option[4]").click() 
		
		gdcl.find_element_by_name("kw").click()
		gdcl.find_element_by_name("kw").clear()
		gdcl.find_element_by_name("kw").send_keys(keyword)
		gdcl.find_element_by_class_name("search_1").submit()
		time.sleep(2)

		return gdcl

	def get_web_gdcl(self,info,starttime,endtime,state):
		weblist = []

		if state is True:
			while True:
				soup = BeautifulSoup(info.page_source)

				web = soup.find_all('a',class_ = "f_l")
				date_all = soup.find_all('a',class_ = "f_r")

				for web_unit in web:
					weblist.append(web_unit['href'])

				for date_unit in date_all:
					self.datelist.append(date_unit.get_text())

				pagedown_index = soup.find('a',text = re.compile(u"下一页»"))
				# info.close()
				pagedown_web = pagedown_index['href']
				pagedown_judge = re.findall("industry=&page=\d+",pagedown_web)

				if len(pagedown_judge) !=0:
					info.get(pagedown_web)
					print("翻1页")
					time.sleep(random.randint(3,6))
				else:
					info.quit()
					print("关闭浏览器")
					break
				
		else:
			while True:
				dateindex = []
				webindex = []
				site = 0

				soup = BeautifulSoup(info.page_source)

				web = soup.find_all('a',class_ = "f_l")
				date_all = soup.find_all('a',class_ = "f_r")
				pagedown_index = soup.find('a',text = re.compile(u"下一页»"))
				pagedown_web = pagedown_index['href']
				pagedown_judge = re.findall("industry=&page=\d+",pagedown_web)

				for web_unit in web:
					webindex.append(web_unit['href'])

				for date_unit in date_all:
					datestrlist = re.findall("\d{4}-\d+-\d+",date_unit.get_text())
					datestr = datestrlist[0]
					dateint = int(datestr[0]+datestr[1]+datestr[2]+datestr[3]+datestr[5]+datestr[6]+datestr[8]+datestr[9]) 
					dateindex.append(dateint)

				for date in dateindex:
					if starttime <= date <= endtime:
						self.datelist.append(date)
						weblist.append(webindex[site])
					else:
						None
					site +=1

				if starttime < dateindex[len(dateindex)-1]:
					if len(pagedown_judge) != 0:
						info.get(pagedown_web)
						print("翻1页")
						time.sleep(random.randint(3,6))
					else:
						info.quit()
						print("关闭浏览器")
						break
				else:
					info.quit()
					print("关闭浏览器")
					break
			
		return weblist

	def get_title_gdcl(self,soup):
		titletag = soup.find('h1',class_ = "title")
		title = titletag.get_text()
		return title

	def get_account_gdcl(self,message_list):
		site = 0
		i = 0
		account = None
		for message in message_list:
			try:
				message.index(u"项目预算金额")
				site = i
			except:
				i +=1
		if site !=0:
			account = re.findall(r"\d*\,?\d*\,?\d+\.?\d*",message_list[site])
			if len(account) !=0:
				return account[0]
			else:
				return "未找到"
		else:
			return "未找到"

	def get_showtime_gdcl(self,message_list):
		site = 0
		i = 0
		showtime = None
		for message in message_list:
			try:
				message.index(u"开标时间")
				site = i
			except:
				i +=1

		if site != 0:
			showtime = re.findall(r"\d+年\d+月\s?\d+日",message_list[site])
			if len(showtime) != 0:
				return showtime[0]
			else:
				return "未找到"
		else:
			return "未找到"


	def get_agent_gdcl(self,message_list):
		
		return "广东采联采购招标有限公司"

	def get_buyer_gdcl(self,message_list):
		site = 0
		i = 0
		buyer = None
		for message in message_list:
			try:
				message.index(u"采购人：")
				site = i
			except:
				try:
					message.index(u"采购单位：")
					site = i
				except:
					try:
						message.index(u"招标人名称：")
						site = i
					except:
						i +=1
		if site != 0:
			return message_list[site]
		else:
			return "未找到"



	def get_detail_gdcl(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
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
			html = urllib.request.urlopen(web,timeout = 30)
			content = html.read()
			html.close()
			"""打开项目网址"""
			message = []

			soup = BeautifulSoup(content)

			message_all = soup.find_all('div', class_="content")
			for message_part in message_all:
				for message_unit in message_part.stripped_strings:
					message.append(message_unit)

			if state1 is True:
				sheet.write(excel_count,5,self.get_title_gdcl(soup))
			else:
				None
			if state2 is True:	
				sheet.write(excel_count,12,self.get_account_gdcl(message))
			else:
				None
			if state3 is True:
				sheet.write(excel_count,0,self.datelist[excel_count-1])
			else:
				None
			if state4 is True:	
				sheet.write(excel_count,3,self.get_agent_gdcl(message))
			else:
				None
			if state5 is True:
				sheet.write(excel_count,1,web)
			else:
				None
			if state6 is True:
				sheet.write(excel_count,11,"测试中标金额")
			else:
				None
			if state7 is True:
				sheet.write(excel_count,7,self.get_showtime_gdcl(message))
			else:
				None
			if state8 is True:
				sheet.write(excel_count,4,self.get_buyer_gdcl(message))
			else:
				None

			excel_count +=1

		wbk.save(filename)

