#!pip install selenium
#!pip install xlwings
#!pip install numpy

import pandas as pd
import time
import os,glob
import xlwings as xw
from random import randint
import re
import numpy as np
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
pd.options.mode.chained_assignment = None



class WebScrapShoes:
    
    def __init__(self):
        
        self.search_result_arr = []
        self.size_col_arr = []
        self.browser = webdriver.Chrome("chromedriver.exe")
        self.browser.get("https://www.dsw.com/en/us/category/mens-shoes/N-1z141hwZ1z141ju")
        self.df_basic_unique = None
        self.df_details_unique = None
        self.df_final = None
        
        
    def scrollToBottom(self):
        
        lenOfPage = self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        match=False
        while(match==False):
                lastCount = lenOfPage
                time.sleep(2)
                lenOfPage = self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
                if lastCount==lenOfPage:
                    match=True
                
    def scrapBasic(self , pages = 42 , products_per_page = 90):
        
        pages = [x for x in range(0, pages*products_per_page, products_per_page)] # as every page contains 90 products

        for page in pages:
            try:
                self.browser.get("https://www.dsw.com/en/us/category/mens-shoes/N-1z141hwZ1z141ju?No="+str(page))
                time.sleep(randint(13,16))
                self.scrollToBottom()

                product_links = self.browser.find_elements_by_class_name("product-tile--link")
                product_ids = self.browser.find_elements_by_class_name("product-tile")
                product_names = self.browser.find_elements_by_class_name("product-tile__name")


                for link,ids,names in zip(product_links,product_ids,product_names ):
                    try:
                        plink = link.get_attribute("href")
                    except Exception:
                        plink = "N/A"
                    try:
                        pid = ids.get_attribute('id')[13:]
                    except Exception:
                        pid = "N/A"
                    try:
                        pname = names.text
                    except Exception:
                        pname = "N/A"

                    self.search_result_arr.append([plink,pid,pname])

            except Exception:
                continue
        
    def exportBasicData(self):
        
        df = pd.DataFrame(self.search_result_arr,columns=['Product URL','Product Id','Product Title'])
        self.df_basic_unique = df.drop_duplicates(subset=['Product URL'], keep='first')
        excel_file_name = "search_result_data.xlsx"
        self.df_basic_unique.to_excel(excel_file_name,index=False)
        print("Excel file of basic data exported")
        
    def scrapDetails(self):
        
        product_link_from_file = pd.read_excel("search_result_data.xlsx")
        count = len(product_link_from_file['Product URL'].values)

        for element in range(0 , count+1):
            try:
                plink = product_link_from_file['Product URL'].values[element]
                self.browser.get(plink)
                time.sleep(randint(5,8))
                carr = []
                sizearr = []
                
                try:
                    pprice = self.browser.find_element_by_id("price").text[1:]
                except Exception as e:
                    pprice = "N/A"

                try:
                    pcolor = self.browser.find_elements_by_class_name("color-swatch-container--price-context")
                    for color in pcolor:
                        details = str(color.find_element_by_tag_name("img").get_attribute("alt"))
                        carr.append(details.split(" ")[-2])
                    pcolor = ", ".join(carr)
                except Exception as e:
                    pcolor = "N/A"

                try:
                    psize = self.browser.find_elements_by_class_name("box-selector")
                    for i in psize[:-1]:
                        if(i.is_enabled()):
                            sizearr.append(i.text)
                    psize = ", ".join(sizearr)       
                except Exception as e:
                    psize = "N/A"

                self.size_col_arr.append([plink,pprice,pcolor,psize])

            except Exception:
                continue
                
    def exportDetailsData(self):
        
        df = pd.DataFrame(self.size_col_arr,columns=['Product URL','Price','Color','Size'])
        self.df_details_unique = df.drop_duplicates(subset=['Product URL'], keep='first')
        excel_file_name = "details_result_data.xlsx"
        self.df_details_unique.to_excel(excel_file_name,index=False)
        print("Excel file of details data exported")
        
                
    def mergeFiles(self):
    
        self.df_final = pd.merge(self.df_basic_unique , self.df_details_unique, on='Product URL',  how='left')
        final_file_name = "final_data.xlsx"
        self.df_final.to_excel(final_file_name,index=False)
        print("Final data exported")


if __name__ == "__main__":

   obj = WebScrapShoes()
   obj.scrapBasic()
   obj.exportBasicData()
   obj.scrapDetails()
   obj.exportDetailsData()
   obj.mergeFiles()
