from selenium import webdriver
from openpyxl.workbook import Workbook
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import datetime

"""
Download the excel sheet and add the path to the Script to do the given instruction
accorcding to the user input.
"""
username = input("Enter the username : ")
password = input("Enter the password : ")

"""
Add Product function,This Function required all the fields to create a new product
you just need to add the product detail which is asked in the Script related to the
product.And after this the Script will automate the creation of the product after
adding it to the excel .
"""

def add_product():
    df = pd.read_excel('All Products.xlsx')

    size=df.shape[0]

    df.drop([0],axis=0,inplace=True)
    df.columns = ['S.no','Name','Category','Sku','Type','Date']
    df1 = df.copy()
    size=df1.shape[0]

    df1.drop([0],axis=0,inplace=True)
    df1.columns = ['S.no','Name','Category','Sku','Type','Date']

    serial_no = input("enter the serial no. :")
    date = datetime.datetime.now()
    title = input("Enter the title :")
    product_data = int(input("Enter select the Product type: 1.Simple and 2.Variable"))
    Vendor_id = int(input("Enter the id : 1.Kazmira LLC , 2.Folium , 3.SilverShadow , 4.CBDHempEx"))
    packaing = int(input('Enter from options :- 1: 1 oz Amber Glass Dropper Bottle , 2: 1 oz White PP Straight Side Jar, 3: 2.5 oz Amber PET Packer Bottle, 4: 1 x 2-Piece Cellulose Hemp Oil Infused Face Mask , 5: 5 oz 150 cc clear PET pill packer bottle,...,14: 2.5 oz Amber PET Packer Bottle'))
    visibility = int(input('Enter from options :- 1:Show , 2:Hide'))
    defalut_price = int(input("Enter the regular price :"))
    sale_price = int(input('Enter the sale price :'))
    sku = int(input('Enter the stock keeping unit :'))
    products_type = input("Enter the value: Oils , Capsules , Topicals , Pets , Edibles")

    df_t=df1.T
    df_t[size] = [serial_no,title,products_type,sku,product_data,date]
    df1 = df_t.T

    for pname in df['Name']:
        for n_name in df1['Name']:
            if pname == n_name:
                print("Product Exits")

            else:
                driver = webdriver.Firefox()
                url = "https://admin.soleralife.com/"
                driver.get(url)
                driver.maximize_window()

                driver.find_element_by_name("""email""").send_keys(username)
                driver.find_element_by_name("""password""").send_keys(password)

                driver.implicitly_wait(1000)
                driver.find_element_by_xpath("""//*[@id="loginform"]/input[2]""").click()
                driver.implicitly_wait(1000)
                driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/a""").click()
                driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/ul/li[1]/a""").click()
                driver.find_element_by_xpath("""//*[@id="producttitle"]""").send_keys(title)

                select1 = Select(driver.find_element_by_id("""product-type"""))
                if product_data == 1:
                    select1.select_by_visible_text('Simple product')
                elif product_data == 2:
                    select1.select_by_visible_text('Variable product')

                select2 = Select(driver.find_element_by_name("""vendorid"""))
                if Vendor_id == 1:
                    select2.select_by_value('5cadf7d854650f1c307468ea')
                elif Vendor_id == 2:
                    select2.select_by_value('5cadf81854650f1c307468eb')
                elif Vendor_id == 3:
                    select2.select_by_value('5cadf82b54650f1c307468ec')
                elif Vendor_id == 4:
                    select2.select_by_value('5cadf83c54650f1c307468ed')
                else:
                    select2.select_by_value('5da3aa208240fe706bf87817')

                select3 = Select(driver.find_element_by_name("""packagingtype"""))
                if packaing == 1:
                    select3.select_by_value('5cc40a2216300f3430662e3d')
                elif packaing == 2:
                    select3.select_by_value('5cc40a2216300f3430662e3e')
                elif packaing == 3:
                    select3.select_by_value('5cc40a2216300f3430662e3f')
                elif packaing == 4:
                    select3.select_by_value('5cc40a2216300f3430662e40')
                elif packaing == 5:
                    select3.select_by_value('5cc40a2216300f3430662e41')
                elif packaing == 6:
                    select3.select_by_value('5cc40a2216300f3430662e42')
                elif packaing == 7:
                    select3.select_by_value('5cc40a2216300f3430662e43')
                elif packaing == 8:
                    select3.select_by_value('5cc40a2216300f3430662e44')
                elif packaing == 9:
                    select3.select_by_value('5cc40a2216300f3430662e45')
                elif packaing == 10:
                    select3.select_by_value('5cc40a2216300f3430662e46')
                elif packaing == 11:
                    select3.select_by_value('5cc40a2216300f3430662e47')
                elif packaing == 12:
                    select3.select_by_value('5cc40a2216300f3430662e48')
                elif packaing == 13:
                    select3.select_by_value('5cc40a2216300f3430662e49')
                elif packaing == 14:
                    select3.select_by_value('5cc40a2216300f3430662e4a')

                select4 = Select(driver.find_element_by_id("""visibility-type"""))
                if visibility == 1:
                    select4.select_by_value('true')
                else:
                    select4.select_by_value('false')

                driver.implicitly_wait(500)
                driver.find_element_by_xpath("""//*[@id="def_regular_price"]""").send_keys(defalut_price)
                driver.find_element_by_xpath("""//*[@id="def_sale_price"]""").send_keys(sale_price)

                driver.find_element_by_xpath("""//*[@id="woocommerce-product-data"]/div/div/ul/li[2]/a""").click()
                driver.implicitly_wait(500)
                driver.find_element_by_xpath("""//*[@id="prod_sku"]""").send_keys(sku)

                driver.implicitly_wait(500)

                driver.find_element_by_xpath("""//*[@id="addproductform"]/div[2]/div[2]/div[2]/span[1]/span[1]/span/ul/li/input""").send_keys(products_type,Keys.ENTER)

                driver.implicitly_wait(1000)
                driver.find_element_by_xpath("""//*[@id="addproductform"]/div[2]/div[1]/div[2]/input""").click()
            df1.to_excel('All Products.xlsx')
            driver.close()

"""
Delete product function , In this you need to enter the name of the product
and follow the instructions to confirm the delete.And the Script will automate the
delete function.
"""

def delete_product():

    conf_del = input("Enter y to confirm and n to cancel :")

    if conf_del == 'y' or conf_del == 'Y':
        del_pro = input("Enter the product name you want to delete : ")

        driver = webdriver.Firefox()
        url = "https://admin.soleralife.com/"
        driver.get(url)
        driver.maximize_window()
        driver.find_element_by_name("""email""").send_keys(username)
        driver.find_element_by_name("""password""").send_keys(password)

        driver.implicitly_wait(1000)
        driver.find_element_by_xpath("""//*[@id="loginform"]/input[2]""").click()
        driver.implicitly_wait(1000)
        driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/a""").click()
        driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/ul/li[2]/a""").click()

        driver.find_element_by_xpath("""//*[@id="DataTables_Table_0_filter"]/label/input""").send_keys(del_pro)
        driver.implicitly_wait(5000)

        driver.find_element_by_xpath("""//*[@id="DataTables_Table_0"]/tbody/tr/td[2]/div/a[2]/i""").click()
        if conf_del == 'y' or conf_del == 'Y':
            driver.find_element_by_xpath("""//*[@id="all_products"]/div[2]/div/div/div[2]/button[2]""").click()
        else:
            driver.find_element_by_xpath("""//*[@id="all_products"]/div[2]/div/div/div[2]/button[1]""").click()
    else:
        print("No delete!")

    driver.close()

"""
Update product, You need to enter the name of the product.
To update the required fields you need to select and
enter the field details you want to change
and Rest leave on the Script to do it's work.
"""

def update_product():
    df = pd.read_excel('All Products.xlsx')

    size=df.shape[0]

    df.drop([0],axis=0,inplace=True)
    df.columns = ['S.no','Name','Category','Sku','Type','Date']
    df1 = df.copy()
    size=df1.shape[0]

    df1.drop([0],axis=0,inplace=True)
    df1.columns = ['S.no','Name','Category','Sku','Type','Date']

    up_pro = input("Enter the product name you want to update : ")
    udf = df1[df1.Name == up_pro]
    ud_f = udf.index
    serial_no=udf['S.no']
    title = udf['Name']
    products_type = udf['Category']
    sku = udf['Sku']
    product_data = udf['Type']
    date = udf['Date']
    defalut_price = int(input("Enter the new regular price"))
    sale_price = int(input('Enter the new sale price'))

    while True:

        more_ele = input("Enter y to select element and n to stop : ")

        if more_ele == 'y' or more_ele == 'Y':
            select_item  = int(input("Enter the elements you want to update : 1.Title ,2.Vendor_id, 3.packaing,  4.visibility , 5.products_type ,6.product_data, 7.sku"))

            if select_item == 1:
                title = input("Enter the new title :")
                udf.loc[ud_f[0],'Name'] = title

            elif select_item == 2:
                Vendor_id = int(input("Enter the id : 1.Kazmira LLC , 2.Folium , 3.SilverShadow , 4.CBDHempEx"))

            elif select_item == 3:
                packaing = int(input('Enter from options :- 1: 1 oz Amber Glass Dropper Bottle , 2: 1 oz White PP Straight Side Jar, 3: 2.5 oz Amber PET Packer Bottle, 4: 1 x 2-Piece Cellulose Hemp Oil Infused Face Mask , 5: 5 oz 150 cc clear PET pill packer bottle,...,14: 2.5 oz Amber PET Packer Bottle'))

            elif select_item == 4:
                visibility = int(input('Enter from options :- 1:Show , 2:Hide'))

            elif select_item == 5:
                products_type = input("Enter the value: Oils , Capsules , Topicals , Pets , Edibles")
                udf.loc[ud_f[0],'Category'] = products_type

            elif select_item == 6:
                product_data = int(input("Enter select the Product type: 1.Simple and 2.Variable"))
                udf.loc[ud_f[0],'Type'] = product_data

            elif select_item == 7:
                sku = int(input('Enter the new stock keeping unit :'))
                udf.loc[ud_f[0],'Sku'] = sku

        else:
            print("Selected Item will be updated shortly!")
            break

    driver = webdriver.Firefox()
    url = "https://admin.soleralife.com/"
    driver.get(url)
    driver.maximize_window()

    driver.find_element_by_name("""email""").send_keys(username)
    driver.find_element_by_name("""password""").send_keys(password)

    driver.implicitly_wait(1000)
    driver.find_element_by_xpath("""//*[@id="loginform"]/input[2]""").click()
    driver.implicitly_wait(1000)
    driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/a""").click()
    driver.find_element_by_xpath("""//*[@id="side-menu"]/li[2]/ul/li[2]/a""").click()

    driver.find_element_by_xpath("""//*[@id="DataTables_Table_0_filter"]/label/input""").send_keys(up_pro)
    driver.implicitly_wait(5000)

    driver.find_element_by_xpath("""//*[@id="DataTables_Table_0"]/tbody/tr/td[2]/div/a[1]""").click()
    driver.find_element_by_xpath("""//*[@id="producttitle"]""").clear()
    driver.find_element_by_xpath("""//*[@id="producttitle"]""").send_keys(title)

    select1 = Select(driver.find_element_by_id("""product-type"""))
    if product_data == 1:
        select1.select_by_visible_text('Simple product')
    elif product_data == 2:
        select1.select_by_visible_text('Variable product')
    else:
        pass

    select2 = Select(driver.find_element_by_name("""vendorid"""))
    if Vendor_id == 1:
        select2.select_by_value('5cadf7d854650f1c307468ea')
    elif Vendor_id == 2:
        select2.select_by_value('5cadf81854650f1c307468eb')
    elif Vendor_id == 3:
        select2.select_by_value('5cadf82b54650f1c307468ec')
    elif Vendor_id == 4:
        select2.select_by_value('5cadf83c54650f1c307468ed')
    elif Vendor_id == 5 :
        select2.select_by_value('5da3aa208240fe706bf87817')
    else:
        pass

    select3 = Select(driver.find_element_by_name("""packagingtype"""))
    if packaing == 1:
        select3.select_by_value('5cc40a2216300f3430662e3d')
    elif packaing == 2:
        select3.select_by_value('5cc40a2216300f3430662e3e')
    elif packaing == 3:
        select3.select_by_value('5cc40a2216300f3430662e3f')
    elif packaing == 4:
        select3.select_by_value('5cc40a2216300f3430662e40')
    elif packaing == 5:
        select3.select_by_value('5cc40a2216300f3430662e41')
    elif packaing == 6:
        select3.select_by_value('5cc40a2216300f3430662e42')
    elif packaing == 7:
        select3.select_by_value('5cc40a2216300f3430662e43')
    elif packaing == 8:
        select3.select_by_value('5cc40a2216300f3430662e44')
    elif packaing == 9:
        select3.select_by_value('5cc40a2216300f3430662e45')
    elif packaing == 10:
        select3.select_by_value('5cc40a2216300f3430662e46')
    elif packaing == 11:
        select3.select_by_value('5cc40a2216300f3430662e47')
    elif packaing == 12:
        select3.select_by_value('5cc40a2216300f3430662e48')
    elif packaing == 13:
        select3.select_by_value('5cc40a2216300f3430662e49')
    elif packaing == 14:
        select3.select_by_value('5cc40a2216300f3430662e4a')
    else:
        pass

    select4 = Select(driver.find_element_by_id("""visibility-type"""))
    if visibility == 1:
        select4.select_by_value('true')
    elif visibility == 2:
        select4.select_by_value('false')
    else:
        pass

    driver.implicitly_wait(500)
    driver.find_element_by_xpath("""//*[@id="def_regular_price"]""").clear()
    driver.find_element_by_xpath("""//*[@id="def_sale_price"]""").clear()
    driver.find_element_by_xpath("""//*[@id="def_regular_price"]""").send_keys(defalut_price)
    driver.find_element_by_xpath("""//*[@id="def_sale_price"]""").send_keys(sale_price)

    driver.find_element_by_xpath("""//*[@id="woocommerce-product-data"]/div/div/ul/li[2]/a""").click()
    driver.implicitly_wait(500)
    driver.find_element_by_xpath("""//*[@id="prod_sku"]""").clear()
    driver.find_element_by_xpath("""//*[@id="prod_sku"]""").send_keys(sku)

    driver.implicitly_wait(500)

    driver.find_element_by_xpath("""//*[@id="addproductform"]/div[2]/div[2]/div[2]/span[1]/span[1]/span/ul/li/input""").send_keys(products_type,Keys.ENTER)

    driver.implicitly_wait(1000)
    driver.find_element_by_xpath("""//*[@id="addproductform"]/div[2]/div[1]/div[2]/input""").click()

    udf.loc[ud_f[0]] = [serial_no,title,products_type,sku,product_data,date]
    driver.close()

"""
Select Work what you want to do with the excel sheet.
"""

select_work = int(input("Enter the work you want to do : 1.Add product , 2.Delete Product , 3.Update Products :"))
if select_work == 1:
    add_product()
elif select_work == 2:
    delete_product()
elif select_work == 3:
    update_product()
