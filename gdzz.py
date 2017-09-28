#-*- coding:utf-8 -*-

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
from selenium import webdriver
import xlwt
import time
import random

class Programme_gdzz():
	datelist = []
	programmenamelist = []
	def get_content_gdzz(self,url,typetext,keyword):

		gdzz = webdriver.PhantomJS()
		gdzz.set_window_size(800,600)
		gdzz.get(url)

		gdzz.find_element_by_id("menu3").click()
		time.sleep(1)
		gdzz.find_element_by_id("keyword").click()
		gdzz.find_element_by_id("keyword").clear()
		gdzz.find_element_by_id("keyword").send_keys(keyword)
		if typetext == "招标类":
			if not gdzz.find_element_by_xpath("//table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/select[2]//option[2]").is_selected():
				gdzz.find_element_by_xpath("//table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/select[2]//option[2]").click()
		else:
			if not gdzz.find_element_by_xpath("//table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/select[2]//option[5]").is_selected():
				gdzz.find_element_by_xpath("//table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/select[2]//option[5]").click()
		gdzz.find_element_by_xpath("//table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/input").click()

		time.sleep(2)
		return gdzz

	def get_web_gdzz(self,info,starttime,endtime,state):
		weblist = []

		if state is True:
			while True:
				dateindex = []
				soup = BeautifulSoup(info.page_source)

				website = soup.find_all('a',href = re.compile("tender.php\?act"))
				dateall = soup.find_all('td',text = re.compile("\d{4}\-\d+\-\d+"))

				for web in website:
					weblist.append(web['href'])
					self.programmenamelist.append(web.get_text())

				for date in dateall:
					self.datelist.append(date.get_text())
					dateindex.append(date.get_text())
				if len(dateindex) < 15:
					break
				else:
					try:
						info.find_element_by_link_text("››").click()
						print("翻一页")
						time.sleep(random.randint(3,5))
					except:
						break		
		else:
			while True:
				webindex = []
				programmenameindex = []
				dateindex = []
				site = 0
				soup = BeautifulSoup(info.page_source)

				website = soup.find_all('a',href = re.compile("tender.php\?act"))
				dateall = soup.find_all('td',text = re.compile("\d{4}\-\d+\-\d+"))

				for web in website:
					webindex.append(web['href'])
					programmenameindex.append(web.get_text())

				for date in dateall:
					datestr = date.get_text()
					dateint = int(datestr[0]+datestr[1]+datestr[2]+datestr[3]+datestr[5]+datestr[6]+datestr[8]+datestr[9])
					dateindex.append(dateint)

				for dateindex_unit in dateindex:
					if starttime <= dateindex_unit <= endtime:
						weblist.append(webindex[site])
						self.programmenamelist.append(programmenameindex[site])
						self.datelist.append(dateindex_unit)
					else:
						None
					site +=1
				if dateindex[len(dateindex)-1] > starttime:
					try:
						info.find_element_by_link_text("››").click()
						print("翻一页")
						time.sleep(random.randint(3,5))
					except:
						break
				else:
					break

		return weblist

	def get_buyer_gdzz(self,message):
		site = 0
		i = 0
		buyer = []

		for message_unit in message:
			try:
				message_unit.index(u"采购人：")
				site = i
			except:
				try:
					message_unit.index(u"采购单位：")
					site = i
				except:
					try:
						message_unit.index(u"招标人：")
						site = i
					except:
						i +=1
		if site !=0:
			buyer.append(message[site])
			return buyer[0]
		else:
			return "未找到"

	# def get_programmename_gdzz(self,message):
	# 	site = 0
	# 	i = 0
	# 	programmename = None

	# 	for message_unit in message:
	# 		try:
	# 			message_unit.index(u"采购项目名称：")
	# 			site = i
	# 		except:
	# 			i +=1
	# 	if site !=0:
	# 		programmename = message[site]
	# 		return programmename
	# 	else:
	# 		return "未找到"

	def get_account_gdzz(self,message):
		site = 0
		i = 0
		account = None

		for message_unit in message:
			try:
				message_unit.index(u"采购项目预算金额")
				site = i
			except:
				try:
					message_unit.index(u"项目预算")
					site = i+1
				except:
					try:
						message_unit.index(u"项目采购预算")
						site = i +1
					except:
						i +=1
		if site != 0:
			account = message[site]
			return account
		else:
			return "未找到"

	def get_agent_gdzz(self,message):
		return "广东志正招标有限公司"

	def get_showtime_gdzz(self,message):
		site = 0
		i = 0
		showtime = None

		for message_unit in message:
			try:
				message_unit.index(u"开标时间：")
				site = i
			except:
				try:
					message_unit.index(u"开标评标时间")
					site = i
				except:
					i +=1
		if site != 0:
			showtime = message[site]
			if len(showtime) <9:
				showtime = message[site +1]
			else:
				None
			return showtime
		else:
			return "未找到"

	# def test(self,WBall):
	# 	site = 0
	# 	for web in WBall:
	# 		print(web,self.programmenamelist[site],self.datelist[site])
	# 		site +=1
	# 	# for web in WBall:
	# 	# 	html = urllib.request.urlopen("http://www.tender.gd.gov.cn/"+web,timeout = 30)
	# 	# 	content = html.read()
	# 	# 	html.close()

	# 		message = []

	# 		soup = BeautifulSoup(content)
			
	# 		message_all = soup.find('td',class_=re.compile("pmain")).stripped_strings
	# 		try:
	# 			for message_all_unit in message_all:
	# 				message.append(message_all_unit)
	# 		except:
	# 			message.append("未找到")
			
		# 	print(self.get_showtime(message))
		
	def get_detail_gdzz(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
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
			html = urllib.request.urlopen("http://www.tender.gd.gov.cn/"+web,timeout = 30)
			content = html.read()
			html.close()

			message = []

			soup = BeautifulSoup(content)
						
			message_all = soup.find('td',class_=re.compile("pmain")).stripped_strings
			try:
				for message_all_unit in message_all:
					message.append(message_all_unit)
			except:
				message.append("未找到")

			if state1 is True:
				sheet.write(excel_count,5,self.programmenamelist[excel_count-1])
			else:
				None
			if state2 is True:
				sheet.write(excel_count,12,self.get_account_gdzz(message))
			else:
				None
			if state3 is True:
				sheet.write(excel_count,0,self.datelist[excel_count-1])
			else:
				None
			if state4 is True:
				sheet.write(excel_count,3,self.get_agent_gdzz(message))
			else:
				None
			if state5 is True:
				sheet.write(excel_count,1,"http://www.tender.gd.gov.cn/"+web)
			else:
				None
			if state6 is True:
				sheet.write(excel_count,11,"")
			else:
				None
			if state7 is True:
				sheet.write(excel_count,7,self.get_showtime_gdzz(message))
			else:
				None
			if state8 is True:
				sheet.write(excel_count,4,self.get_buyer_gdzz(message))
			else:
				None

			excel_count +=1

		wbk.save(filename)		

