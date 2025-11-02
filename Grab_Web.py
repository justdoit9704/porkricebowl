import ShareWork
import iSelenium
import os
import ShareWork as shareWork
import time
import math
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime, timedelta


def get_latest_file(download_dir, timeout=30):
    """
    等待下载完成后，返回目录中最新的文件名
    """
    seconds = 0
    while seconds < timeout:
        files = os.listdir(download_dir)
        if files:
            files = [os.path.join(download_dir, f) for f in files]
            latest_file = max(files, key=os.path.getctime)
            # 检查文件是否还在下载（部分浏览器下载文件会以 .crdownload/.part 后缀存在）
            if not latest_file.endswith((".crdownload", ".part")):
                return latest_file
        time.sleep(0.5)
        seconds += 0.5

    if latest_file is None:
        raise FileNotFoundError("下载文件未找到或超时")

    return latest_file


def start(curpath, userInfo, date_range, date_list_yymmdd):

    # configfloder = os.path.join(curpath, 'config')
    user_path = os.path.join(os.path.join(curpath, 'resources'), userInfo[2])
    ShareWork.createFolder(user_path)
    main_json = os.path.join(curpath, 'Main_Config.json')
    JSON = shareWork.getJsonData(main_json)

    download_path = JSON["Path"]["Downloads"].format(user_path)
    ShareWork.createFolder(download_path)
    pdf_path = JSON["Path"]["PDF"].format(user_path)
    ShareWork.createFolder(pdf_path)
    excel_path = JSON["Path"]["Excel"].format(user_path)
    ShareWork.createFolder(excel_path)

    min_date = min(date_list_yymmdd)
    max_date = max(date_list_yymmdd)

    print("Start Download Grab PDF File......")

    driver = iSelenium.driver(download_path)
    driver.get(JSON["Link"]["Grab"].format(curpath))

    # Username
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="Username"]')
    driver.find_element(By.XPATH,'//*[@id="Username"]').send_keys(userInfo[0])

    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="root"]/section/div/div[2]/div[1]/div/form/div/div[3]/div/div/span/button').click()

    # Password
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="password"]')
    driver.find_element(By.XPATH,'//*[@id="password"]').send_keys(userInfo[1])

    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="root"]/section/div/div[2]/div[1]/div/form/button').click()

    # close notic
    iSelenium.presenceElemWait(driver, "xpath", '/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div[3]/button[1]')
    driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div[3]/button[1]').click()

    # order page
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="root"]/div/div/div[1]/aside/div[1]/div[1]/ul/li[3]/div')
    driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div[1]/aside/div[1]/div[1]/ul/li[3]/div').click()

    # history
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-tab-history"]')
    driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-tab-history"]').click()

    # Complete status
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-history"]/div/div[1]/div[2]/div/div/span[2]')

    recp_status_elem = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-history"]/div/div[1]/div[2]/div/div/span[2]')
    recp_status_elem.click()

    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="food"]/div[2]/div/div/div[2]/div/div/div/div[2]/div')
    driver.find_element(By.XPATH, '//*[@id="food"]/div[2]/div/div/div[2]/div/div/div/div[2]/div').click()

    for date in date_range:
        # 日历
        print(date + "\nProcessing......")
        iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-history"]/div/div[1]/div[1]/div/div')
        driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-history"]/div/div[1]/div[1]/div/div').click()

        date_elem = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-history"]/div/div[1]/div[1]/div/div/input')
        date_elem.send_keys(Keys.CONTROL, 'a')  # 全选
        date_elem.send_keys(date)
        date_elem.send_keys(Keys.ENTER)

        # 检查今天的order数量
        iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-history"]/div/div[2]/div[3]/div/div/div[2]')
        order_num_xpath = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-history"]/div/div[2]/div[3]/div/div/div[2]')
        print("Order Number: " + order_num_xpath.text.strip())

        time.sleep(1)
        grab_order = int(order_num_xpath.text.strip())
        if grab_order > 10:
            loop_number = math.ceil(grab_order / 10)
        else:
            loop_number = 1

        for i in range(loop_number):
            i+=1
            if i == 1:
                # 选择前10账单
                iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-history"]/div/div[3]/div/div/div/div/div[1]/table/thead/tr/th[1]/div/label/span/input')
                all_tick = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-history"]/div/div[3]/div/div/div/div/div[1]/table/thead/tr/th[1]/div/label/span/input')
                time.sleep(2)
                if all_tick.is_enabled():
                    all_tick.click()
                else:
                    continue

            else:
                # 选择后10账单
                all_tick.click()
                time.sleep(2)
                all_tick.click()
                time.sleep(2)

                for n in range(10):
                    n = (i-1)*10+n
                    try:
                        # iSelenium.presenceElemWait(driver, "xpath",
                        #                            f'//tr[@data-row-key="{n}"]//child::td//child::label')
                        row = driver.find_element(By.XPATH, f'//tr[@data-row-key="{n}"]//child::td//child::label')
                        row.click()
                    except NoSuchElementException:
                        break

            # 下载账单
            iSelenium.presenceElemWait(driver, "xpath",
                                           '//*[@id="rc-tabs-0-panel-history"]/div/div[4]/div/div[1]/div[1]/div[2]/div/div/button')
            driver.find_element(By.XPATH,
                                    '//*[@id="rc-tabs-0-panel-history"]/div/div[4]/div/div[1]/div[1]/div[2]/div/div/button').click()

            dt = datetime.strptime(date, "%a, %d %b %Y")
            file_path = os.path.join(download_path, "ReceiptMultipleOrder.pdf")
            final_pdf_path = os.path.join(pdf_path, "{}_{}.{}".format(dt.strftime("%Y%m%d"), i, "pdf"))

            shareWork.checkfileExist(file_path, 60)
            shareWork.moveFile(file_path, final_pdf_path)
        print("Completed......")

    print("Start Download Grab Excel File......")

    # Finace Page
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="root"]/div/div/div[1]/div/aside/div[1]/div[1]/ul/li[5]/div')
    driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div[1]/div/aside/div[1]/div[1]/ul/li[5]/div').click()
    # //*[@id="root"]/div/div/div[1]/aside/div[1]/div[1]/ul/li[5]/

    # 先等开始日期的xpath
    iSelenium.presenceElemWait(driver, "xpath",
                               '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/div/span[1]/div/div/div[1]/input')

    # 确保没有弹窗出来
    driver.refresh()

    # time.sleep(20)
    # Next button
    if iSelenium.check_exists_by_xpath(driver, '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[1]/div/div/div[2]/div/button/span'):
        driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[1]/div/div/div[2]/div/button/span').click()
    else:
        pass

    # 开始日期
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/div/span[1]/div/div/div[1]/input')
    start_date_elem = driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/div/span[1]/div/div/div[1]/input')
    start_date_elem.click()
    start_date_elem.send_keys(Keys.CONTROL, 'a')  # 全选
    # time.sleep(2)
    start_date_elem.send_keys(min_date)
    start_date_elem.send_keys(Keys.ENTER)

    time.sleep(2)

    # 结束日期
    iSelenium.presenceElemWait(driver, "xpath",
                               '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/div/span[1]/div/div/div[3]/input')
    end_date_elem = driver.find_element(By.XPATH,
                                    '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/div/span[1]/div/div/div[3]/input')
    end_date_elem.click()
    end_date_elem.send_keys(Keys.CONTROL, 'a')  # 全选
    end_date_elem.send_keys(max_date)
    end_date_elem.send_keys(Keys.ENTER)

    # 下载
    time.sleep(2)
    iSelenium.presenceElemWait(driver, "xpath", '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/span/button')
    driver.find_element(By.XPATH, '//*[@id="rc-tabs-0-panel-transactions"]/div/div/div[2]/div/span/button').click()

    # print(download_path)
    latest_file = get_latest_file(download_path, 60)
    min_date = (datetime.strptime(min_date, "%Y-%m-%d")).strftime("%Y%m%d")
    max_date = (datetime.strptime(max_date, "%Y-%m-%d")).strftime("%Y%m%d")
    final_excel_path = os.path.join(excel_path,"{}_{}.{}".format(min_date, max_date, "csv"))
    shareWork.moveFile(latest_file, final_excel_path)
    print("Completed......")

    driver.quit()


# start(os.getcwd(), ["dayday.manager","Aa113322*"], ["Fri, 15 Aug 2025"], ["2025-08-02", "2025-08-03"])