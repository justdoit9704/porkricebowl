import json
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains as AC, ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException


def driver(path):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920x1080")

    # chrome_options.add_argument("--disable-popup-posting")
    # chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--disable-notifications")
    # chrome_options.add_argument("--no-sandbox")
    # chrome_options.add_argument("--verbose")
    # chrome_options.add_argument("--disable-extensions")
    # chrome_options.add_argument("--disable-web-security")
    # chrome_options.add_argument("--disable-logging")
    # chrome_options.add_argument("--disable-popup-posting")

    appState = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }

    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": r"{}".format(path),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False,
        "printing.print_preview_sticky_settings.appState": json.dumps(appState),
        "savefile.default_directory": r"{}".format(path),  # 指定保存路径
    })
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    # print(driverPath + "//chromedriver.exe")
    # chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--kiosk-printing")

    return webdriver.Chrome(options=chrome_options)


def wait(driver):
    return WebDriverWait(driver, 10)


def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


def presenceElemWait(driver, eltype, elem):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]
    el=""
    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((elemTypeArray[elemIndex], elem))
            )
            break
        except Exception as e:
            # print("{0}: {1}".format(elem,e))
            print("Waiting......\n" + elem)

    return


def PresenceCondWait(driver, elem, attr, value):

    while True:
        try:
            print(driver.find_element(By.XPATH,elem).get_attribute(attr))
            if driver.find_element(By.XPATH,elem).get_attribute(attr) == value:
                break
        except Exception as e:
            print("{0}: {1}".format(elem,e))

    return


def clickElemWait(driver, eltype, elem):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]

    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((elemTypeArray[elemIndex], elem))
            )
            break
        except Exception as e:
            print("{0}: {1}".format(elem,e))

    return


def clickElemConditionWait(driver, eltype, elem):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]

    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((elemTypeArray[elemIndex], elem))
            )
            break
        except Exception as e:
            print("{0}: {1}".format(elem, e))

    return


def elemInvisible(driver, eltype, elem):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]

    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.invisibility_of_element((elemTypeArray[elemIndex], elem))
            )
            break
        except Exception as e:
            print("{0}: {1}".format(elem, e))

    return


def elemVisible(driver, eltype, elem):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]

    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((elemTypeArray[elemIndex], elem))
            )
            break
        except Exception as e:
            print("{0}: {1}".format(elem, e))

    return


def scrollToEnd(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")


def scrollToUp(driver):
    driver.execute_script("window.scrollTo(0, 0);")


def executeScript_elem_click(driver, eltype, elem, type=None):
    elemNameArray = ['id', 'class', 'xpath', 'tag', 'css', 'link', 'name']
    elemTypeArray = [By.ID, By.CLASS_NAME, By.XPATH, By.TAG_NAME, By.CSS_SELECTOR, By.LINK_TEXT, By.NAME]

    elemIndex = 0

    for e in elemNameArray:
        if eltype.lower() == e.lower():
            break

        elemIndex += 1

    script = "document.getElement"
    if type != None:
        script += "s{}({}).click()".format(elemTypeArray[elemIndex], elem)
    else:
        script += "{}({}).click()".format(elemTypeArray[elemIndex], elem)

    driver.execute_script(script)

    return

def executeScript_changeValue(driver, elem, value):
    print('document.querySelector("{}").value = {}'.format(elem, value))
    driver.execute_script('document.querySelector("{}").value = {}'.format(elem, value))

    return


def executeScript_query_element_click(driver, elem, value):
    print('document.querySelector("{}").click()'.format(elem, value))
    driver.execute_script('document.querySelector("{}").click()'.format(elem, value))

    return


def executeScript_query_allelement_click(driver, elem, index):
    print('document.querySelectorAll("{}")[{}].click()'.format(elem, index))
    driver.execute_script('document.querySelectorAll("{}")[{}].click()'.format(elem, index))

    return


def executeScript_query_element_change(driver, eltype, value, elem):
    print('arguments[0].setAttribute("{}","{}")'.format(elem, value))
    driver.execute_script('arguments[0].setAttribute("{}","{}")'.format(eltype, value), elem)

    return


def executeNewWindow(driver, link):
    print("window.open('{}');".format(link))
    driver.execute_script("window.open('{}');".format(link))

    return

def switchLatestWindow(driver):
    windows = driver.window_handles
    driver.switch_to.window(windows[len(windows) - 1])


def actionChain_click(driver, elem):
    AC(driver).move_to_element(elem).click(elem).perform()


def actionMove(driver, elem):
    AC(driver).move_to_element(elem).perform()


def switchIframeWIndex(driver, index):
    driver.switch_to.frame(driver.find_element(By.TAG_NAME,"iframe")[index])


def escapeKeys(driver):
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()