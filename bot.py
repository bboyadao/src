from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
login_url = "https://accounts.shopify.com/store-login"
admin_url = "https://kourse-com.myshopify.com/admin"
shop_address = "https://kourse-com.myshopify.com"
email = "bradevans87@hotmail.com"
pass_w = "T3stdeveloperp4ss"

options = webdriver.ChromeOptions()
options.add_argument(
    "user-data-dir=/Users/xiaoma/Library/Application Support/Google/Chrome/")
driver = webdriver.Chrome(chrome_options=options)
driver.get(login_url)
shop_elem = driver.find_element_by_xpath('//*[@id="shop_domain"]')

# shop_elem.send_keys(shop_address)
_click1 = driver.find_element_by_xpath(
    '//*[@id="body-content"]/div[1]/div[2]/div/form/button')
_click1.click()


email_elem = driver.find_element_by_xpath('//*[@id="account_email"]')
email_elem.send_keys(email)
_click2 = driver.find_element_by_xpath('//*[@id="js-login-form"]/form/button')
_click2.click()

time.sleep(10)
pass_elem = driver.find_element_by_xpath('//*[@id="account_password"]')
pass_elem.send_keys(pass_w)

# _click3 = driver.find_element_by_xpath('//*[@id="login_form"]/button')
# _click3.click()
input("Press Enter to continue once capatcha has been successfully completed...")
print(driver.current_url)

#
#
# WebDriverWait(driver, 10).until(EC.url_changes(changed_url))
