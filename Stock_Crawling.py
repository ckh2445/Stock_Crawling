import requests
from bs4 import BeautifulSoup
import openpyxl

class Stock_Crawling():
    def __init__(self):
        self.res =  requests.get("https://finance.naver.com/sise/")
        self.soup = BeautifulSoup(self.res.content,"html.parser")
        
    def get_name(self):
        return self.soup.select("div.rgt > ul.lst_major > li > a")
    
    def get_price(self):
        return self.soup.select("div.rgt > ul.lst_major > li > span")
    
    def get_state(self):
        return self.soup.select("div.rgt > ul.lst_major > li > em")

    def save(self,filename:str,lists:list):
        self.excel_file = openpyxl.Workbook()
        self.excel_sheet = self.excel_file.active
        self.excel_sheet.column_dimensions['A'].width = 12
        self.excel_sheet.column_dimensions['B'].width = 15
        self.excel_sheet.column_dimensions['C'].width = 20
        self.title = ["종목","가격","상승 및 하락"]
        
        self.excel_sheet.append(self.title)
        self.excel_sheet.title = filename
        
        for item in lists:
            self.excel_sheet.append(item)
            
        self.excel_file.save(filename)
        self.excel_file.close()
if __name__ == "__main__":
    Stock = Stock_Crawling()
    lists = list()
    
    name = Stock.get_name()
    price = Stock.get_price()
    state = Stock.get_state()
    
    for idx,item in enumerate(name):
        lists.append([item.get_text().strip(),price[idx].get_text()+"원",state[idx].get_text()])
        
    Stock.save("Stock.xlsx",lists)