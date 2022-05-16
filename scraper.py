from abc import ABC, abstractmethod
from concurrent.futures import ProcessPoolExecutor
from time import sleep
from random import randint

from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd


SAVE_PATH = Path(__file__).parent / 'scraped/'


class AbstractScraper(ABC):

    @abstractmethod
    def __enter__(self):
        pass

    @abstractmethod
    def __exit__(self, *exc_info):
        pass

    """закрыть и переключиться на предыдушую страницу"""
    def switch_to_page(self, page, close=True):
        if close:
            self.driver.close()
        self.driver.switch_to.window(page)
        sleep(randint(1, 3))


class CategoryScraper(AbstractScraper):

    def __init__(self, name, link) -> None:
        self.name = name
        self.link = link

    @staticmethod
    def get_save_path(cat_name):
        return SAVE_PATH / f'{cat_name}.xlsx'

    def __enter__(self):
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        save_path = self.get_save_path(self.name)
        self.writer = pd.ExcelWriter(str(save_path), engine='openpyxl')
        return self
    
    def __exit__(self, *exc_info):
        self.writer.close()
        self.driver.quit()

    def run(self):
        self.driver.get(self.link)

        sub_main_page = self.driver.window_handles[0]
        rubrics_categories = self.driver.find_element(By.CLASS_NAME, 'rubricsCategories')
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
            self.driver.execute_script("window.open('');")
            page_of_subcategories = self.driver.window_handles[1]
            self.driver.switch_to.window(page_of_subcategories)
            self.driver.get(f"{sub_cat_link}?pagenumber=1&pagesize=100")
            sleep(randint(1, 3))

            """пагинация"""
            page = 1
            while True:
                try:
                    pagination = self.driver.find_element(By.CSS_SELECTOR, 'ul.pagination')
                except:
                    pagination = None   

                if pagination is None:
                    break
                elif pagination and pagination.text != '':
                    self.scrape_subcategory(page_of_subcategories, sheet_name)
                    page += 1 
                    pagination = self.driver.find_element(By.CSS_SELECTOR, 'ul.pagination')
                    last_page = int(pagination.find_elements(By.TAG_NAME, 'li')[-1].text)
                    if page < last_page:
                        self.driver.get(f"{sub_cat_link}?pagenumber={page}&pagesize=100")
                        self.driver.switch_to.window(page_of_subcategories)
                        continue
                    elif page == last_page:
                        self.driver.get(f"{sub_cat_link}?pagenumber={page}&pagesize=100")
                        self.scrape_subcategory(page_of_subcategories, sheet_name)
                        break
                elif pagination.text == '':
                    self.scrape_subcategory(page_of_subcategories, sheet_name)
                    break  
            self.switch_to_page(sub_main_page)

    """скрапер данных подкатегории"""
    def scrape_subcategory(self, page_of_subcategories, sheet_name):
        organizations = self.driver.find_elements(By.CSS_SELECTOR, 'a.organizationName.blueText')
        subcategory_data = []
        for organization in organizations:
            try:
                data = self.scrape_organization(organization)
            except:
                continue
            subcategory_data.append(data)
            self.switch_to_page(page_of_subcategories)
        df = pd.DataFrame(subcategory_data)
        try:
            df.to_excel(self.writer, sheet_name=f"{sheet_name}")
        except:
            return
        else:
            self.writer.save()
    
    def scrape_organization(self, organization):
        link = organization.get_attribute('href')
        self.driver.execute_script("window.open('');")
        parent = self.driver.window_handles[3]
        self.driver.switch_to.window(parent)
        self.driver.get(link)
        sleep(randint(1, 3))
        main_table = self.driver.find_element(By.CSS_SELECTOR, 'div.organizationPage')
        ps = main_table.find_elements(By.TAG_NAME, 'p')
        title = (self.driver.find_element(By.CSS_SELECTOR, 'h1.text25.mt20').text).replace(' - КОНТАКТЫ, АДРЕС, ТЕЛЕФОН', '')
        phone_number = (self.driver.find_element(By.CSS_SELECTOR, 'p.text16.lh23').text).replace(' ', '')
        email = ''
        web_site = ''
        legal_name = ''
        office_hours = ''
        address = (self.driver.find_element(By.CSS_SELECTOR, 'p.address').text).replace('Адрес: ', '')

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

        return {'Организация': title, 'Юридическое название': legal_name, 'Брендовое название': brand_name, 'Контактные номера': phone_number, 'e-mail': email, 'Веб-страница': web_site, 'Адрес': address, "Режим работы": office_hours}



class Scraper(AbstractScraper):
    
    def __init__(self) -> None:
        self.executor = ProcessPoolExecutor(max_workers=5)

    def __enter__(self):
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        return self

    def __exit__(self, *exc_info):
        self.executor.shutdown()
        self.driver.quit()
    
    def run(self):
        self.driver.get('https://www.yellowpages.uz/')
        categories = self.driver.find_elements(By.CLASS_NAME, 'media-heading')
        print(f"Found total {len(categories)} categories")
        for category in categories:
            if not CategoryScraper.get_save_path(category.text).exists():
                self.scrape_category(category)
        self.driver.quit() 

    def scrape_category(self, category):
        print(f"Starting category: {category.text}")
        excel_name = category.text
        cat_link = (category.find_element(By.TAG_NAME, 'a')).get_attribute('href')
        with CategoryScraper(excel_name, cat_link) as category_scraper:
            category_scraper.run()


if __name__ == "__main__":
    with Scraper() as scraper:
        scraper.run()
