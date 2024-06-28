import time
import os
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import ddddocr
from selenium.common.exceptions import NoSuchElementException, ElementNotSelectableException
from win32com.client import Dispatch
import wget
import zipfile


def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    version = parser.GetFileVersion(filename)
    return version


def ocr(file):
    # 验证码识别
    ocr = ddddocr.DdddOcr()
    res = ocr.classification(file)
    return res


def download_edgedriver(version):
    wget.download("https://msedgedriver.azureedge.net/" + version + "/edgedriver_win64.zip", "./edgedriver_win64.zip")
    file = zipfile.ZipFile("./edgedriver_win64.zip")
    file.extractall("./")
    file.close()
    os.rename("./msedgedriver.exe", "./MicrosoftWebDriver.exe")
    os.remove("./edgedriver_win64.zip")


if __name__ == '__main__':
    browser = None
    bro = input("选择你所使用的浏览器：1：firefox 2：chrome 3：edge\n")
    if bro == '1':
        from selenium.webdriver.firefox.options import Options
        try:
            options=Options()
            options.binary_location=r"C:/Program Files/Mozilla Firefox/firefox.exe"
            browser = webdriver.Firefox(options=options)
        except Exception as E:
            print(E)
            print("trying firefox nightly")
            try:
                options=Options()
                options.binary_location=r"C:/Program Files/Firefox Nightly/firefox.exe"
                browser = webdriver.Firefox(options=options)
            except Exception:
                print("failed")
                exit(-1)
    elif bro == '2':
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ['enable-automation'])
        options.add_argument("--disable-blink-features=AutomationControlled")
        browser = webdriver.Chrome(options=options)
    elif bro == '3':
        edgepath = r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
        if os.path.exists(edgepath):
            edgeversion = get_version_via_com(edgepath)
            print("Your msedge version is ", edgeversion)
            driverpath = r"./MicrosoftWebDriver.exe"
            if os.path.exists(driverpath):
                driverversion = get_version_via_com(driverpath)
                print("Your msedgedriver version is ", driverversion)
                if driverversion != edgeversion:
                    os.remove("./MicrosoftWebDriver.exe")
                    download_edgedriver(edgeversion)
            else:
                download_edgedriver(edgeversion)
        options = webdriver.EdgeOptions()
        options.add_experimental_option("excludeSwitches", ['enable-automation'])
        options.add_argument("--disable-blink-features=AutomationControlled")
        browser = webdriver.Edge(options=options, executable_path=".")
    browser.get("http://cas.sysu.edu.cn/cas/login?service=https%3A%2F%2Fjwxt.sysu.edu.cn%2Fjwxt%2Fapi%2Fsso%2Fcas%2Flogin%3Fpattern%3Dstudent-login&locale=zh_CN")
    browser.find_element(By.XPATH, '//img[@id="captchaImg"]').screenshot("t.png")
    f = open("t.png", "rb+")
    data = f.read()
    browser.find_element(By.XPATH, "//input[@placeholder='验证码']").send_keys(ocr(data))
    input("登录，按回车继续")
    browser.find_element(By.XPATH, "//div[text()='我的评教']").click()
    window_handles = browser.window_handles
    browser.switch_to.window(window_handles[1])
    time.sleep(5)
    while True:
        try:
            try:
                browser.find_element(By.XPATH, "//div[contains(@class,'cz-bo')]/div[1]//button").click()
            except NoSuchElementException:
                print("评教完成\n")
                input("按回车结束\n")
                browser.quit()
                quit(0)
            except ElementNotSelectableException:
                print("评教完成，有部分课程无法完成评教，请手动评教\n")
                input("按回车结束\n")
                browser.quit()
                quit(0)
            time.sleep(0.5)
            elements = len(browser.find_elements(By.XPATH, "//tr[contains(@class,'ant-table-row-level-0')]"))
            for i in range(1, elements):
                browser.find_element(By.XPATH, "//tr[contains(@class,'ant-table-row-level-0')][" + str(i) + "]//input").send_keys(browser.find_element(By.XPATH, "//tr[contains(@class,'ant-table-row-level-0')][" + str(i) + "]/td[3]").text)
            browser.find_element(By.XPATH, "//span[text()='提 交']/..").click()
            try:
                WebDriverWait(browser, 10, 0.5).until_not(EC.presence_of_element_located((By.XPATH, "//i[contains(@class,'anticon-loading')]")))
            except Exception as e:
                print(type(e), "::", e)
            try:
                browser.find_element(By.XPATH, "//input[@placeholder='请输入验证码']").click()
            except NoSuchElementException:
                print("成功，无验证码")
            else:
                ele = browser.find_element(By.XPATH, '//img[@alt="验证码"]')
                ele.screenshot("t.png")
                f = open("t.png", "rb+")
                data = f.read()
                browser.find_element(By.XPATH, "//input[@placeholder='请输入验证码']").send_keys(ocr(data))
                browser.find_element(By.XPATH, "//span[text()='提 交']/..").click()
                WebDriverWait(browser, 10, 0.5).until_not(EC.presence_of_element_located((By.XPATH, "//i[contains(@class,'anticon-loading')]")))
                try:
                    browser.find_element(By.XPATH, "//input[@placeholder='请输入验证码']")
                except NoSuchElementException:
                    print("成功，有验证码")
        except Exception as e:
            print(e)
        finally:
            browser.refresh()