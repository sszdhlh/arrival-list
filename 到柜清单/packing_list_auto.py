import webbrowser
from selenium.webdriver.common.keys import Keys
import time
from selenium import webdriver

# create a new webbrowser
driver = driver = webdriver.Chrome()

# open the target url
driver.get(
    "https://go.cin7.com/Cloud/ShoppingCartAdmin/Orders/OrdersList.aspx?idWebSite=12877&idCustomerAppsLink=526109")

# wait 2 seconds
time.sleep(2)

# 定位输入框并输入数据
username_elem = driver.find_element_by_id("https://www.cin7.com/ accounts@trsports.com.au")  # 请更改为实际的元素ID
password_elem = driver.find_element_by_id("71424709Tr")  # 请更改为实际的元素ID


