import os

from time import sleep
from random import randint

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd


class Scrape():
    def scrape_web_site(self):
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get('https://www.yellowpages.uz/')
        categories = driver.find_elements(By.CLASS_NAME, 'media-heading')
        for category in categories:
            excel_name = category.text
            cat_link = (category.find_element(By.TAG_NAME, 'a')).get_attribute('href')
            driver.execute_script(f"window.open('{cat_link}', 'new_window')")

            """новый excel файл с названием категории"""
            writer = pd.ExcelWriter(f'{excel_name}.xlsx', engine='openpyxl')

            main_page = driver.window_handles[0]
            sub_main_page = driver.window_handles[1]
            driver.switch_to.window(sub_main_page)
            sleep(randint(1, 5))


            rubrics_categories = driver.find_element(By.CLASS_NAME, 'rubricsCategories')
            sub_cats = rubrics_categories.find_elements(By.CSS_SELECTOR, 'a.text-bold.darkText')
            for sub_cat in sub_cats:
                """новый excel sheet с названием подкатегории"""
                sheet_name = sub_cat.text
                sheet_name = sheet_name.replace('/', '')
                sheet_name = sheet_name.replace('*', '')
                sheet_name = sheet_name.replace('?', '')
                sheet_name = sheet_name.replace(':', '')
                sheet_name = sheet_name.replace('[', '')
                sheet_name = sheet_name.replace(']', '')


                sub_cat_link = sub_cat.get_attribute('href')
                driver.execute_script("window.open('');")
                page_of_subcategories = driver.window_handles[2]
                driver.switch_to.window(page_of_subcategories)
                driver.get(f"{sub_cat_link}?pagenumber=1&pagesize=100")
                sleep(randint(5, 15))

                """пагинация"""
                page = 1
                while True:
                    try:
                        pagination = driver.find_element(By.CSS_SELECTOR, 'ul.pagination')
                    except:
                        pagination = None   

                    if pagination is None:
                        break
                    elif pagination and pagination.text != '':
                        scrape_data(driver, page_of_subcategories, writer, excel_name, sheet_name)
                        page += 1 
                        pagination = driver.find_element(By.CSS_SELECTOR, 'ul.pagination')
                        last_page = int(pagination.find_elements(By.TAG_NAME, 'li')[-1].text)
                        if page < last_page:
                            driver.get(f"{sub_cat_link}?pagenumber={page}&pagesize=100")
                            driver.switch_to.window(page_of_subcategories)
                            continue
                        elif page == last_page:
                            driver.get(f"{sub_cat_link}?pagenumber={page}&pagesize=100")
                            scrape_data(driver, page_of_subcategories, writer, excel_name, sheet_name)
                            break
                    elif pagination.text == '':
                        scrape_data(driver, page_of_subcategories, writer, excel_name, sheet_name)
                        break  
                close_window(driver, sub_main_page)  
            writer.close()      
            close_window(driver, main_page)            
        driver.quit() 


"""закрыть и переключиться на предыдушую страницу"""
def close_window(driver, page):
    driver.close()
    driver.switch_to.window(page)
    sleep(randint(1, 5))


"""скрапер данных данных подкатегории"""
def scrape_data(driver, page_of_subcategories, writer, excel_name, sheet_name):
    organizations = driver.find_elements(By.CSS_SELECTOR, 'a.organizationName.blueText')
    for organization in organizations:
        link = organization.get_attribute('href')
        driver.execute_script("window.open('');")
        parent = driver.window_handles[3]
        driver.switch_to.window(parent)
        driver.get(link)
        sleep(randint(5, 15))
        main_table = driver.find_element(By.CSS_SELECTOR, 'div.organizationPage')
        ps = main_table.find_elements(By.TAG_NAME, 'p')
        title = (driver.find_element(By.CSS_SELECTOR, 'h1.text25.mt20').text).replace(' - КОНТАКТЫ, АДРЕС, ТЕЛЕФОН', '')
        phone_number = (driver.find_element(By.CSS_SELECTOR, 'p.text16.lh23').text).replace(' ', '')
        email = ''
        web_site = ''
        legal_name = ''
        brand_anme = ''
        office_hours = ''
        address = (driver.find_element(By.CSS_SELECTOR, 'p.address').text).replace('Адрес: ', '')

        for p in ps:
            try:
                email_ref = p.find_element(By.XPATH, "//*[contains(text(), 'E-mail: ')]")                   
            except:
                email_ref = ''
        for p in ps:        
            try:
                web_site_ref = p.find_element(By.XPATH, "//img[contains(@src,'/Content/images/Website.png')]/following-sibling::a")                 
            except:
                web_site_ref = ''
        for p in ps:        
            try:
                legal_name_ref = p.find_element(By.XPATH, "//*[contains(text(), 'Юридическое название: ')]")                   
            except:
                legal_name_ref = ''  
        for p in ps:        
            try:
                brand_name_ref = p.find_element(By.XPATH, "//*[contains(text(), 'Брендовое название: ')]")
            except:
                brand_name_ref = ''   
        for p in ps:        
            try:
                office_hours_ref = p.find_element(By.XPATH, "//*[contains(text(), 'Часы работы: ')]")
            except:
                office_hours_ref = ''               
        for p in ps:                        
            if email_ref and email_ref.text in p.text:
                email = (p.text).replace('E-mail: ', '')
            elif web_site_ref and  web_site_ref.text in p.text: 
                web_site = p.text
            elif legal_name_ref and legal_name_ref.text in p.text:
                legal_name = (p.text).replace('Юридическое название: ', '')
            elif brand_name_ref and brand_name_ref.text in p.text:
                brand_name = (p.text).replace('Брендовое название: ', '')
            elif office_hours_ref and office_hours_ref.text in p.text:
                office_hours = (p.text).replace('Часы работы: ', '')                  

        df = pd.DataFrame([{'Организация': title, 'Юридическое название': legal_name, 'Брендовое название': brand_name, 'Контактные номера': phone_number, 'e-mail': email, 'Веб-страница': web_site, 'Адрес': address, "Режим работы": office_hours}], index=pd.RangeIndex(start=1, name='index'))
        if os.path.getsize(f"{excel_name}.xlsx") == 0 or sheet_name not in writer.sheets:
            df.to_excel(writer, sheet_name=f"{sheet_name}")
        else: 
            df.to_excel(writer, header=False, index=True, startrow=writer.sheets[f"{sheet_name}"].max_row, sheet_name=f"{sheet_name}")
        ws = writer.sheets[sheet_name]
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))    
        for col, value in dims.items():
            ws.column_dimensions[col].width = value
        writer.save()
        close_window(driver, page_of_subcategories)
             

data = Scrape()



if __name__ == "__main__":
    data.scrape_web_site()