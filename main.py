import pandas as pd
from selenium import webdriver
import time

'''Read Excel Spreadsheet'''
df_fyi = pd.read_excel('equity.xlsx')

sid_list = []

count = 80 # for keeping track of where in SID list we are

'''Put SIDs in list'''
for index, row in df_fyi.iterrows():
    sid_list.append(row["STUDENT_ID"])

browser = webdriver.Chrome('C:/chromedriver/chromedriver')

browser.get('https://caccl-mendocino.alma.exlibrisgroup.com/ng;u=%2Fmng%2Faction%2Fhome.do%3FngHome%3Dtrue')


'''Enter Username and Password - Then Login'''
user_name = browser.find_element_by_xpath("//input[@id='username']")
user_name.send_keys("admin")

pass_word = browser.find_element_by_xpath("//input[@id='password']")
pass_word.send_keys("") # This must be added before running

browser.find_element_by_xpath("//input[@value='Login']").click()

time.sleep(5)

while count < len(sid_list):
    SID_Code = sid_list[count]  # send individual SID from sid_list to be used

    '''Open Desired Page on Alma'''
    browser.get('https://caccl-mendocino.alma.exlibrisgroup.com/ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dloan.fulfillment_checkout.xml&pageViewMode%3DEdit&operation%3DLOAD&pageBean.currentUrl%3DxmlFileName%253Dloan.fulfillment_checkout.xml%2526pageViewMode%253DEdit%2526operation%253DLOAD%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&menuKey%3Dcom.exlibris.dps.wrk.general.menu.Fulfillment.Advanced.Settings.checkout&pageBean.securityHashToken%3D-7053974269830552792;ng=true')

    time.sleep(3)

    '''Send SID from list'''
    browser.find_element_by_xpath("//input[@name='pageBean.displayNameOfUserOrUserIdendifier']").send_keys(SID_Code)

    '''Navigate through menus'''
    time.sleep(2)

    browser.find_element_by_name("page.buttonsPointers[0].operation").click()
    count += 1
    time.sleep(2)

    try:
        browser.find_element_by_link_text("Edit Notes").click()
        print("boy howdy")

        time.sleep(2)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[3]/table/tbody/tr/td[9]/div/button").click()
        time.sleep(1)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[3]/table/tbody/tr/td[9]/div/ul/li[1]/a").click()
        time.sleep(1)
        browser.find_element_by_name("pageBean.updateNote.userNote.note").click()
        time.sleep(1)
        browser.find_element_by_name("pageBean.updateNote.userNote.note").send_keys(" and EQUITY " + "Fall 2021")
        time.sleep(1)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[2]/div[2]/div[1]/button").click()
    except (BaseException):
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[4]/div/div[4]/div[2]/table/tbody/tr/td[3]/a").click()
        time.sleep(2)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[1]/div[2]/div/ul/li[2]/div/div/button").click()
        time.sleep(1)
        browser.find_element_by_name("pageBean.newNote.userNote.note").click()
        time.sleep(1)
        browser.find_element_by_name("pageBean.newNote.userNote.note").send_keys("Equity Fall 2021")
        time.sleep(1)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[1]/div[2]/div/ul/li[2]/div/div/ul/li/span/div/div[3]/div/div/div[1]/span/button").click()
        time.sleep(1)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[1]/div[2]/div/ul/li[2]/div/div/ul/li/span/div/div[3]/div/div/div[1]/ul/ul/li[5]").click()
        time.sleep(1)
        browser.find_element_by_xpath("/html/body/div[3]/base-root/main-layout/ex-main-layout/div[2]/div[1]/page-wrapper/div/div/div/div/form/div[4]/div/div[5]/div/div[2]/div[1]/div[2]/div/ul/li[2]/div/div/ul/li/span/div/div[8]/div/div[2]/span/button").click()
        time.sleep(1)
