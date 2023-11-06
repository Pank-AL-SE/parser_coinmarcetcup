import datetime
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from tkinter import *
from tkinter import ttk
from functools import partial
import openpyxl

def parsing(data_all_coins):
    browser = webdriver.Chrome()
    browser.get("https://coinmarketcap.com/")
    for i in range(1, 12):
        browser.execute_script("window.scrollTo(0, %d);" % (800 * i))

    browser.set_page_load_timeout(15)
    for number in range(50):
        table = browser.find_element(By.XPATH, '//tbody/tr[%d]' % (number + 1))
        name = table.find_element(By.XPATH, "td[3]//p")
        symbol = table.find_element(By.XPATH, "td[3]/div/a/div/div/div/p")
        price = table.find_element(By.XPATH, "td[4]")
        market_cap = table.find_element(By.XPATH, "td[8]")
        data_one_coin = []
        data_one_coin.append(number + 1)
        data_one_coin.append(name.text)
        data_one_coin.append(symbol.text)
        data_one_coin.append(price.text)
        data_one_coin.append(market_cap.text)
        data_all_coins.append(data_one_coin)
        
    browser.close()
    return data_all_coins

def btn1f():
    table.delete(*table.get_children()) 
    data_all_coins = []
    data_all_coins = parsing(data_all_coins=data_all_coins)
    data = [(data_all_coins[i][0],data_all_coins[i][1],
             data_all_coins[i][2],data_all_coins[i][3],
             data_all_coins[i][4]) for i in range(50)] 
        
    for inf in data :
        table.insert("",END,values=inf)
    
 
    today = datetime.datetime.today()
    lbl.config(text='Last update: '+today.strftime("%Y-%m-%d %H.%M.%S"))
    return data

def btn2f(data_all_coins):
    def DELETE():
        table.delete(*table.get_children())
    def btnfunc(data_all_coins=data_all_coins):        
        str = pole.get()
        for i in range(len(data_all_coins)):
            if data_all_coins[i][1]==str:     
                z = [data_all_coins[i][1],data_all_coins[i][2],data_all_coins[i][3],data_all_coins[i][4]]        
                table.insert("",END,values=z)
                break  

    search = Tk()
    search.title('Searching')
    search.resizable(width=0, height=0)
    lbl = Label(search, text="Введите название криптовалюты:")    
    btn = Button(search, text="Search", command=btnfunc)
    btn1 = Button(search, text="CLEAR", command=DELETE)
    pole = Entry(search)    
    columns = ('nam','symb','pric','market_kup')
    table = ttk.Treeview(search, columns=columns,show='headings')    
    table.heading('nam', text='Name')
    table.heading('symb', text='Symbol')
    table.heading('pric', text='Price')
    table.heading('market_kup', text='Market Kup')    
    lbl.grid(row=0,column=0,sticky="e")
    pole.grid(row=0,column=1,sticky="w")
    btn.grid(row=0,column=2,sticky="w")
    btn1.grid(row=0,column=3,sticky="w")
    table.grid(row=1, columnspan=4)
    search.mainloop()     

def btn3f(data):
    wb = openpyxl.Workbook()
    wb.create_sheet(index = 0, title = "Save Data")
    sheet = wb["Save Data"]
    sheet["A1"] = "Name"
    sheet["B1"] = "Symbol"
    sheet["C1"] = "Price"
    sheet["D1"] = "Market Cap"
    for i in range(50):
        sheet["A"+str(i+2)] = str(data[i][1])
        sheet["B"+str(i+2)] = str(data[i][2])
        sheet["C"+str(i+2)] = str(data[i][3])
        sheet["D"+str(i+2)] = str(data[i][4])
        wb.save("Save.xlsx")
    return 

data_all_coins = []
data_all_coins = parsing(data_all_coins=data_all_coins)
data = [(data_all_coins[i][0],data_all_coins[i][1],
        data_all_coins[i][2],data_all_coins[i][3],
        data_all_coins[i][4]) for i in range(50)]

window = Tk()
window.title("Parsing CoinMarketCap")
window.resizable(width=0, height=0)

btn1 = Button(window, text='Update', background="yellow", command=btn1f)
btn2 = Button(window, text='Search', background="yellow", command=partial(btn2f,data))
btn3 = Button(window, text='Save', background="yellow", command=partial(btn3f,data))
lbl = Label(text = 'Update data!',background="green")

columns = ('num','nam','symb','pric','marketkup')
table = ttk.Treeview(window,columns=columns,show='headings')
table.heading('num', text='№')
table.heading('nam', text='Name')
table.heading('symb', text='Symbol')
table.heading('pric', text='Price')
table.heading('marketkup', text='Market Kup')



btn1.grid(row=0,column=0)
btn2.grid(row=0,column=1)
btn3.grid(row=0,column=2)
table.grid(row = 1,columnspan=3, sticky='nsew')
lbl.grid(row=2, columnspan=3,sticky="nsew")

table.delete(*table.get_children()) 

        
for inf in data :
    table.insert("",END,values=inf)
    
 
today = datetime.datetime.today()
lbl.config(text='Last update: '+today.strftime("%Y-%m-%d %H.%M.%S"))


window.mainloop()