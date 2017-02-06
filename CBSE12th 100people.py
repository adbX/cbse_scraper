import time
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Firefox()

i = 9107821
wb = Workbook()
ws = wb.active
ws.append(["Roll","Name","Total","Eng","Mat","Phy","Chem","IDK"])

for j in range(75):

   driver.get("http://cbseresults.nic.in/class12/cbse1216.htm")
   roll_no = driver.find_element_by_name('regno')
   roll_no.send_keys(str(i))

   school_code = driver.find_element_by_name('schcode')
   school_code.send_keys('72625')

   submit_btn = driver.find_element_by_xpath("/html/body/table[4]/tbody/tr[4]/td/input[1]")
   submit_btn.click()

   for tr_1 in driver.find_elements_by_xpath("//table[1]//tr[2]//td[5]"):
      mark1 = float(tr_1.text)

   for tr_2 in driver.find_elements_by_xpath("//table[1]//tr[3]//td[5]"):
      mark2 = float(tr_2.text)

   for tr_3 in driver.find_elements_by_xpath("//table[1]//tr[4]//td[5]"):
      mark3 = float(tr_3.text)

   for tr_4 in driver.find_elements_by_xpath("//table[1]//tr[5]//td[5]"):
      mark4 = float(tr_4.text)

   for tr_5 in driver.find_elements_by_xpath("//table[1]//tr[6]//td[5]"):
      mark5 = float(tr_5.text)

   for name in driver.find_elements_by_xpath("/html/body/div[2]/table[2]/tbody/tr[2]/td[2]"):
      name1 = name.text

   row = [i,name1,((mark1+mark2+mark3+mark4+mark5)/5), mark1, mark2, mark3, mark4, mark5]
   ws.append(row)

   time.sleep(.01)
   i+=1
   print j
   wb.save("marks.xlsx")

print "ALL DONE AND SAVED!!!"

