from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import openpyxl

###This project will
def main():
    driver = webdriver.Chrome(executable_path=r"C:\Drivers\chromedriver\chromedriver.exe")
    url = "https://www.ebay.com/"
    a = 1
    card_list = ["yugioh"]
    print("This program will search for cards on ebay by lowest price first(on the site) and "
          "wth a 'Buy it now' option set\n")
    c = ''
    while c != 'n':
        card_list.append(input("Please enter which card you would like to look up on ebay: "))
        print(f"Total cards: {card_list}")
        c = input("\nContinue? Type y/n: ")
    print("These cards will now be searched for")

    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    search_button = driver.find_element(By.ID, "gh-ac")
    ###Must have a search result before applying (buy it now)/(lowest price) options. In this case I have card_list[0]
    ### as a placeholder. Output will then be formatted and sent to excel file###
    search_button.send_keys(card_list[0])
    search_button.send_keys(Keys.RETURN)
    time.sleep(1)
    best_match_button = driver.find_element(By.XPATH, '/ html / body / div[5] / div[4] / div[1] / div[2] / div[1] / div[3] / div[1] / div / span / button')
    best_match_button.click()
    lowest_price_button = driver.find_element(By.XPATH,'/html/body/div[5]/div[4]/div[1]/div[2]/div[1]/div[3]/div[1]/div/span/span/ul/li[4]/a')
    lowest_price_button.click()
    buy_it_now = driver.find_element(By.XPATH,'/html/body/div[5]/div[4]/div[1]/div[2]/div[1]/div[2]/div[1]/div/ul/li[4]/a')
    buy_it_now.click()
    path = r"C:\Users\Elliot Mollman\Documents\WebSrapingEbay.xlsx"
    try:
        for item in card_list:
            search_button = driver.find_element(By.ID, "gh-ac")
            time.sleep(2)
            search_button.clear()
            search_button.send_keys(card_list[a])
            search_button.send_keys(Keys.RETURN)
            time.sleep(1)
            price = driver.find_elements(By.CLASS_NAME, "s-item__price")
            price_list = []
            for value in price:
                price_list.append(value.text)
            print(price_list)
            df = pd.DataFrame(data=price_list, columns=[card_list[a]])
            print(df)
            a += 1
            with pd.ExcelWriter(path, mode="a", engine="openpyxl") as writer:
                writer.book = openpyxl.load_workbook(path)
                df.to_excel(writer, sheet_name='sheet1')
    except TimeoutException:
        print("Loading took too much time!")
if __name__ =="__main__":
    main()