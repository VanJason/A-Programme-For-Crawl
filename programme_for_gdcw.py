#-*- coding:utf-8 -*-

"""
广东政府采购网
"""

import re
import urllib
from urllib import request
from bs4 import BeautifulSoup
from selenium import webdriver
import xlwt
import time

class Programme_gdcw():
	datelist = []

	def get_content_gdcw(self,url,typetext,keyword,buyer):
		"""doc.
	    网站打开方法
	    返回网站
		"""
		gdcw = webdriver.PhantomJS()
		gdcw.set_window_size(800,600)
		gdcw.get(url)
		gdcw.find_element_by_link_text(u"信息公告").click()
		gdcw.find_element_by_link_text(typetext).click()
		gdcw.find_element_by_xpath(u"(//a[contains(text(),'全部')])[3]").click()
		gdcw.find_element_by_name("title").clear()
		gdcw.find_element_by_name("title").send_keys(keyword)		
		gdcw.find_element_by_xpath('//input[@name="purchaserOrgName"]').clear()
		gdcw.find_element_by_xpath('//input[@name="purchaserOrgName"]').send_keys(buyer)
		gdcw.find_element_by_css_selector("input.n_m_s_but").click()

		return gdcw

	def get_Web_gdcw(self,info,starttime,endtime,state):
		"""
		获取检索出的项目的链接
		实现翻页
		"""
		weblist=[]
		if state is True:
			while True:
				soup = BeautifulSoup(info.page_source)
				website = soup.find_all(href=re.compile("showNotice"))

				for web in website:
					gotoweb = web['href']
					weblist.append(gotoweb)
				time.sleep(3)	

				nextpage = soup.find('a',text = re.compile(u"下一页"))
				try:
					judgenextpage = nextpage['href']
					info.execute_script('var page=document.getElementsByName("pageIndex");var count=Number(page[0].value)+1;turnOverPage(count);')
					time.sleep(3)
				except:
					break
			
		else:
			while True:
				soup = BeautifulSoup(info.page_source)
				nextpage = soup.find('a',text = re.compile(u"下一页"))
				dateall = soup.find_all('em', text = re.compile("\d{4}-\d+-\d+\s\d+:\d+"))
				for date in dateall:
					datestr  = date.get_text()
					dateint = int(datestr[0]+datestr[1]+datestr[2]+datestr[3]+datestr[5]+datestr[6]+datestr[8]+datestr[9])
					self.datelist.append(dateint)
				if endtime < self.datelist[len(self.datelist)-1]:
					try:
						judgenextpage = nextpage['href']
						info.execute_script('var page=document.getElementsByName("pageIndex");var count=Number(page[0].value)+1;turnOverPage(count);')
						time.sleep(5)
					except:
						break
				else:
					website = soup.find_all(href=re.compile("showNotice"))

					for web in website:
						gotoweb = web['href']
						weblist.append(gotoweb)
					time.sleep(3)

					if starttime < self.datelist[len(self.datelist)-1]:
						try:
							judgenextpage = nextpage['href']
							info.execute_script('var page=document.getElementsByName("pageIndex");var count=Number(page[0].value)+1;turnOverPage(count);')
							time.sleep(5)
						except:
							break
					else:
						break
				

		return weblist


	def get_title_gdcw(self,soup):
		"""
		获取项目名称的方法
		"""
		title = soup.find('div', class_="zw_c_c_title")

		return title.get_text()

	def get_beginningtime_gdcw(self,soup):
		"""
		获取公布时间的方法
		"""
		programme_head = soup.find('div', class_="zw_c_c_qx")

		head = []
		beginningtime = []

		try:
			for he in programme_head.stripped_strings:
				head.append(he)

		except:
			head.append(None)

		for head_unit in head:
			try:
				head_unit.index(u"发布日期")
				beginningtime.append(head_unit)

			except:
				None

		if len(beginningtime) !=0:
			Btime = beginningtime[0]
			return Btime[5:15]
		else:
			Btime = "无法获取"
			return Btime

	def get_agentcompany_gdcw(self,soup):
		"""
		获取代理机构的方法
		"""
		programme_head = soup.find('div', class_="zw_c_c_qx")

		head = []
		agentcompany = []

		try:
			for he in programme_head.stripped_strings:
				head.append(he)

		except:
			head.append(None)

		for head_unit in head:
			try:
				head_unit.index(u"代理机构")
				agentcompany.append(head_unit)

			except:
				None
		if len(agentcompany) != 0:
			agentname = agentcompany[0]
			return agentname[5:]
		else:
			agentname = "无法获取"
			return agentname

	def get_account_gdcw(self,soup):
		"""
		获取项目预算的方法
		"""
		detail = soup.find('div', class_="zw_c_c_cont")

		message = []
		account = []

		try:
			for detail_unit in detail.stripped_strings:
				message.append(detail_unit)

		except:
			message.append(None)

		for message_unit in message:
			try:
				account_number = message_unit.index(u"、采购项目预算金额")
				account.append(message_unit)

			except:
				try:
					account_number = message_unit.index(u"采购预算")
					account.append(message_unit)
				except:
					None
		if len(account) != 0:
			money = account[0]
			return money[14:]
		else:
			money = 0
			return money 

	def get_showtime_gdcw(self,soup):
		"""
		获取开标时间
		"""
		detail = soup.find('div', class_="zw_c_c_cont")

		message = []
		time = []

		i = 0

		try:
			for detail_unit in detail.stripped_strings:
				message.append(detail_unit)

		except:
			message.append(None)

		for message_unit in message:
			try:
				time_unit = message_unit.index(u"、开标时间")
				time.append(message_unit)
				site = i

			except:
				i += 1

		if len(time) != 0:
			showtime = message[site+1] + message[site+2] + message[site+3] + message[site+4]+ message[site+5]+ message[site+6]
		else:
			showtime = "无法获取"

		return showtime

	def get_buyer_gdcw(self,soup):
		"""
		获取采购人的方法
		"""
		detail = soup.find('div', class_="zw_c_c_cont")

		message = []
		buyer_all = []

		try:
			for detail_unit in detail.stripped_strings:
				message.append(detail_unit)

		except:
			message.append(None)

		for message_unit in message:
			try:
				buyer_unit = message_unit.index(u"采购人：")
				buyer_all.append(message_unit)

			except:
				None
		if len(buyer_all) != 0:
			buyer = buyer_all[0]
			return buyer[7:]
		else:
			buyer = "无法获取"		
			return buyer 
	def get_money_gdcw(self,soup):
		tbodylist = []
		tdlist = []
		countlist = []
		money = 0
		i = 2

		tagtable = soup.find_all('table')
		for tag in tagtable:
			tbodylist.append(tag)

		tagtd = tbodylist[0].find_all('td')
		for td in tagtd:
			tdlist.append(td.get_text())
		tdlen = len(tdlist)/6
		while i <= tdlen:
			countlist.append(tdlist[i*6-1])
			i+=1

		for count in countlist:
			countstr = re.findall(r"\d*\,?\d*\,?\d+\.?\d*",count)

			for strunit in countstr:
				if strunit is not "":
					moneystr = strunit.replace(',','')
					money = money + float(moneystr)	
				
		return money


	def get_programme_gdcw(self,WBall,filename,state1,state2,state3,state4,state5,state6,state7,state8):
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

		i = 1
		for web in WBall:
			try:
				html = urllib.request.urlopen('http://www.gdgpo.gov.cn' + web,timeout = 30)
				content = html.read().decode('utf-8')
				html.close()
			except:
				html = urllib.request.urlopen('http://www.gdgpo.gov.cn' + web, timeout = 30)
				content = html.read().decode('utf-8')
				html.close()
			"""打开项目网址"""

			soup = BeautifulSoup(content)

			if state1 is True:
				sheet.write(i,5,self.get_title_gdcw(soup))
			else:
				None
			if state2 is True:	
				sheet.write(i,12,self.get_account_gdcw(soup))
			else:
				None
			if state3 is True:
				sheet.write(i,0,self.get_beginningtime_gdcw(soup))
			else:
				None
			if state4 is True:	
				sheet.write(i,3,self.get_agentcompany_gdcw(soup))
			else:
				None
			if state5 is True:
				sheet.write(i,1,"http://www.gdgpo.gov.cn" + web)
			else:
				None
			if state6 is True:
				sheet.write(i,11,self.get_money_gdcw(soup))
			else:
				None
			if state7 is True:
				sheet.write(i,7,self.get_showtime_gdcw(soup))
			else:
				None
			if state8 is True:
				sheet.write(i,4,self.get_buyer_gdcw(soup))
			else:
				None

			i +=1

		wbk.save(filename)
		

	# info = get_content('http://www.gdgpo.gov.cn/')
	# WBall= get_Web(info)
	# print(get_programme(WBall))

# p = Programme_gdcw()
# info = p.get_content_gdcw("http://www.gdgpo.gov.cn/","采购公告","","广州市第十二人民医院")
# WBall = p.get_Web_gdcw(info,None,None,True)
# p.get_programme_gdcw(WBall,"gdcwtest.xls",True,True,True,True,True,True,True,True)



