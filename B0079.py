# selenium4
# 待機処理の書き方 参考>>https://www.teru2teru.com/python/selenium/do-not-time-sleep/#google_vignette

print(123)

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.chrome.service import Service # 1) Serviceのインポート
from selenium.webdriver.common.by import By
# from selenium.webdriver.support import expected_conditions as EC

# from selenium.webdriver.support.select import Select
# from selenium.webdriver.common.action_chains import ActionChains

import time
import os

# ドライバーを自動でインストールしてくれる
os.environ["https_proxy"] = "http://proxya10-2.intra.lixil.co.jp:8080"

# driver = webdriver.Chrome(ChromeDriverManager().install())    # Selenium3の場合
driver = webdriver.Chrome(service= Service(ChromeDriverManager().install()))    # Selenium4の場合

# 最大の読み込み時間を設定 今回は最大30秒待機できるようにする
wait = WebDriverWait(driver=driver, timeout=30)


# url = 'https://ds-note.net/coding/selenium_operation/' #対象のURLを指定
# res = driver.get(url)

# dropdown = driver.find_element(By.ID, "dropdown")

# #ドロップダウンを選択する
# select = Select(dropdown) #ドロップダウンメニューが選択された状態

# #値を選択する
# select.select_by_index(2)  # 3番目のoptionタグを選択状態に



# js_code = '''document.getElementById("selectAggregateItem").querySelector("option[value='ET']").selected = true;'''


# # SR来館システムURL
# url = "https://a1681.corp.lixil.lan/srintegrationsystem/ui/"
# # ログインID・PW
# id = "kyoko1.kato@lixil.com"
# pw = "BorDer06231313"

# try:
#     driver.get(url) # URLを開く
#     wait.until(EC.presence_of_all_elements_located) # 要素が全て検出できるまで待機する
#     time.sleep(3)

#     # ■ログインIDの入力 ---
#     # elems = driver.find_elements(by=By.CLASS_NAME, value="input") # Selenium4の場合
#     # elems[0].click()

#     # h1_text = driver.find_element_by_tag_name("h1").text # Selenium3の場合
#     input_email = driver.find_element(by=By.ID, value="i0116") # Selenium4の場合
#     input_email.send_keys(id)

#     elems = driver.find_elements(by=By.ID, value="idSIButton9") # Selenium4の場合
#     elems[0].click()

#     wait.until(EC.presence_of_all_elements_located) # 要素が全て検出できるまで待機する
#     time.sleep(1)
#     # ---

#     # ■パスワードの入力 ---
#     input_email = driver.find_element(by=By.ID, value="i0118") # Selenium4の場合
#     input_email.send_keys(pw)

#     elems = driver.find_elements(by=By.ID, value="idSIButton9") # Selenium4の場合
#     elems[0].click()

#     wait.until(EC.presence_of_all_elements_located) # 要素が全て検出できるまで待機する
#     time.sleep(1)
#     # ---

#     # ■ サインイン状態確認 ---
#     elems = driver.find_elements(by=By.ID, value="idBtn_Back") # Selenium4の場合
#     elems[0].click()

#     wait.until(EC.presence_of_all_elements_located) # 要素が全て検出できるまで待機する
#     time.sleep(5)
#     # ---


#     driver.maximize_window()
#     # driver.find_element(by=By.ID, value="clearBtn").click()

#     js_code = 'document.getElementById("selectAggregateItem").options[8].classList.add("selected");'
#     driver.execute_script(js_code)
#     selector = "#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div > ul"
#     js_code = f'document.querySelector({selector}).li[8].classList.add("selected");'
#     driver.execute_script(js_code)
#     #form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div > ul



    # select_dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'selectAggregateItem')))
    # select_dropdown.click()
    # select_dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'selectAggregateItem')))
    # select = Select(select_dropdown)
    # select.select_by_index(1)


#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div


    # cls_elem = driver.find_element(by=By.CLASS_NAME, value ='col-xs')
    # ul_elems = driver.find_elements(by=By.ID, value="selectAggregateItem")

    # selecter = "#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div"
    # js_code = f"document.querySelector({selecter}).click();"
    # driver.execute_script(js_code)

    # js_code = 'document.querySelector("#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div > ul > li:nth-child(9)").selected = true;'
    # driver.execute_script(js_code)
    # wait.until(EC.presence_of_all_elements_located) # 要素が全て検出できるまで待機する
    # time.sleep(3)

# document.querySelector("#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div > ul > li:nth-child(9)").selected = true;
#form1 > div.Absolute-Center > div:nth-child(1) > div.panel-body > div:nth-child(1) > div:nth-child(8) > div > div > ul > li:nth-child(9)


    # driver.execute_script("arguments[5].click();", ul_elem)


    # li_elems = ul_elem.find_elements(by=By.TAG_NAME, value="option")


    # select = Select(ul_elem)
    # select.select_by_index(2)

    # li_elems = ul_elem.find_elements(by=By.TAG_NAME, value="option")

    # # print(li_elements.text)
    # li_elems[68].click()
    # for li_elem in li_elems:
    #     text = li_elem.text
    #     print(text)
        # if text == "ＬＨＴ特需開発":
        #     li_element.click()
        #     break

  # ここまで


    # print("「" + h1_text + "」のURLを開きました。")
#     print("終了します。")
# # エラーが発生した時はエラーメッセージを吐き出す。
# except Exception as e:
#     print(e)
#     print("エラーが発生しました。")

# 最後にドライバーを終了する
# finally:
driver.close()
driver.quit()