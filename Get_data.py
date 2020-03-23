import time
import random2
from Utility import XLUtility
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome(executable_path="C:\\Users\\Prashant\\PycharmProjects\\GTMatrix\\Driver\\chromedriver.exe")
#driver = webdriver.Firefox(executable_path="C:\Users\Prashant\PycharmProjects\GTMatrix\Driver\geckodriver.exe")

driver.get("https://gtmetrix.com/")
driver.maximize_window()
driver.implicitly_wait(100)
path = "C:\\Users\\Prashant\\Desktop\\GTmatrixx\\Utility\\VS_Data.xlsx"


def login():
    driver.find_element_by_id("li-email").send_keys("nipam@munro-tailoring.com")
    time.sleep(3)
    driver.find_element_by_id("li-password").send_keys("Munro@123!")
    driver.find_element_by_xpath("//*[@id='menu-site-nav']/div[2]/div[1]/form/div[4]/button").click()
    time.sleep(5)




def open_Page(url1,ran_loc):
    driver.find_element_by_xpath("/html/body/div[1]/main/article/form/div[3]/a").click()
    Select(driver.find_element_by_name('region')).select_by_visible_text(ran_loc)
    driver.find_element_by_xpath("/html/body/div[1]/main/article/form/div[3]/a").click()
    driver.find_element_by_name("url").send_keys(url1)
    driver.find_element_by_xpath("/html/body/div[1]/main/article/form/div[1]/div[2]/div/button").click()
    WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.CLASS_NAME, "report-details")))


def check():
    row_count = XLUtility.getRowCount(path, "Sheet2")
    return row_count

def PageType(i):
    if (i == "https://www.ateliermunro.com/" or i == "https://eu.suitsupply.com/en_GB/home" or i == "https://www.oger.nl/" or i == "https://www.harryrosen.com/en/" ):
        PT = "HomePage"
    elif (i == "https://ateliermunro.com/custom-suits" or i == "https://apac.suitsupply.com/en/overview-suits" or i == "https://www.oger.nl/en/online-shop/all-clothing/heren/suits" or i == "https://www.harryrosen.com/en/clothing/tailored-clothing/c/suits"):
        PT = "Suits Overview"
    elif (i == "https://ateliermunro.com/custom-shirts"or i == "https://apac.suitsupply.com/en/overview-shirts"or i == "https://www.oger.nl/en/online-shop/all-clothing/heren/shirts"or i == "https://www.harryrosen.com/en/clothing/c/dress-shirts" ):
        PT = "Shirts Overview"
    elif (i == "https://ateliermunro.com/Navy-Wool-Silk-Linen-Stripe-Suit-MS0343-custom-suits/styledetail/4218?OrderCreation=true"or i == "https://apac.suitsupply.com/on/demandware.store/Sites-APAC-Site/en/Configurator-Show?type=suit"):
        PT = "Configrator"
    elif (i == "https://ateliermunro.com/bookfitting"or i == "https://oger-by-atelier-munro.appointlet.com/"or i == "https://www.harryrosen.com/en/atelier-munro-tailoring"):
        PT = "Book Appointment"
    return PT


def get_result(url1,rowNumber,PT):

    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 11, PT)
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 4, url1)

    CD = datetime.now()
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 1, CD.strftime("%Y-%m-%d"))
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 2, CD.strftime("%H:%M:%S"))

    Region = driver.find_element_by_xpath("//div[@class='report-details-info']/div[2]/div")
    Rg = Region.text
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 3, Rg)

    VL =[]
    Value = driver.find_element_by_xpath("//div[@class='report-details-info']/div[3]/div")
    VL.append(Value.text)
    Tot_val = VL[0].split()

    Device = Tot_val[1]
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 9, Device)
    # XLUtility.writeData(path, "Sheet2", rowNumber+2, 4, Device[1])

    Browser = Tot_val[0]
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 10, Browser)
    #XLUtility.writeData(path, "Sheet2", rowNumber + 1, 5, Browser)

    Page_score = driver.find_element_by_xpath("//div[@class='box clear'] /div[1]/span/span")
    PS = list(Page_score.text)
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 5, PS[1]+PS[2])

    Load_time = driver.find_element_by_xpath("//div[@class='report-page-detail']/span")
    LT = (Load_time.text).split('s')
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 6, LT[0])

    Page_size = driver.find_element_by_xpath("//div[@class='report-page-detail report-page-detail-size']/span")
    Pgs = (Page_size.text).split('MB')
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 7, Pgs[0])

    Requests = driver.find_element_by_xpath("//div[@class='report-page-detail report-page-detail-requests']/span")
    XLUtility.writeData(path, "Sheet2", rowNumber + 1, 8, Requests.text)





driver.find_element_by_id("user-nav-login").click()
login()
loc = ["Vancouver, Canada", "London, UK", "Sydney, Australia", "Dallas, USA", "Mumbai, India", "São Paulo, Brazil","Hong Kong, China"]
ran_loc = random2.choice(loc)
url_list = ["https://www.ateliermunro.com/","https://ateliermunro.com/bookfitting","https://ateliermunro.com/custom-suits","https://ateliermunro.com/custom-shirts","https://ateliermunro.com/Navy-Wool-Silk-Linen-Stripe-Suit-MS0343-custom-suits/styledetail/4218?OrderCreation=true",
           "https://eu.suitsupply.com/en_GB/home","https://apac.suitsupply.com/en/overview-shirts","https://apac.suitsupply.com/en/overview-suits","https://apac.suitsupply.com/on/demandware.store/Sites-APAC-Site/en/Configurator-Show?type=suit",
          "https://www.oger.nl/","https://oger-by-atelier-munro.appointlet.com/","https://www.oger.nl/en/online-shop/all-clothing/heren/suits","https://www.oger.nl/en/online-shop/all-clothing/heren/shirts",
           "https://www.harryrosen.com/en/","https://www.harryrosen.com/en/atelier-munro-tailoring","https://www.harryrosen.com/en/clothing/tailored-clothing/c/suits","https://www.harryrosen.com/en/clothing/c/dress-shirts"]

for i in url_list:
    open_Page(i,ran_loc)
    row_count = check()
    PType=PageType(i)
    get_result(i,row_count,PType)
    time.sleep(5)
    driver.find_element_by_xpath("html/body/div[1]/header/div/nav/ul/li[1]/a/i").click()


driver.close()