import xml.etree.ElementTree as ET
import xlsxwriter
import datetime
from os.path import basename
import os
import itertools
from datetime import date
import shutil

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#blok odczytu configa ze ścieżkami:
#keys - lista powiązań nazwa grupy z kodami dostawców
keys = []
#rec_groups - lista powiązań nazwa grupy z adresami e-mail
rec_groups = []
config_read_list = []
file_path = "C:/XML_to_CSV/config.cfg"
file = open(file_path, "r", encoding='utf8')
lines = file.readlines()

for lin in lines:
    config_read_list.append(lin.replace("\n", ""))
file.close()

source_path = (config_read_list[1].split(" = ")[1]).replace("\\", "/")
picked_directories = list(str((config_read_list[2].split(" = ")[1]).replace("\\", "/")).split(","))
direction_path = (config_read_list[3].split(" = ")[1]).replace("\\", "/")
param_GA = int(config_read_list[4].split(" = ")[1])
for step in range(0, param_GA, 1):
    rec_groups.append((config_read_list[5 + step]).split(" = "))
key_from_file = (config_read_list[5 + param_GA]).split(" = ")[1]
key_from_file = key_from_file.split(";")
for el in key_from_file:
    keys.append(el.split(","))

source_path_c = source_path
direction_path_c = direction_path
picked_directories_c = picked_directories

# czyszczenie folderu po poprzednim uruchomieniu skryptu
try:
    shutil.rmtree(direction_path_c) 
    os.mkdir(direction_path_c)
except:
    os.mkdir(direction_path_c)      

#metoda przetwarzająca
def parser(directory):
    directory_path = (source_path_c + directory.strip() + "/PROGNOZA")
    print("\ndirectory_path:", directory_path)
    workbook = xlsxwriter.Workbook(direction_path_c + "SL_" + str(directory).strip() + ".xlsx")
    worksheet = workbook.add_worksheet()

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    string_format = workbook.add_format({'num_format': '@'})
    int_format = workbook.add_format({'num_format': '#'})

    excel_columns = ("Position", "DocOrderNumber", "DocOrderDate", "BuyerItemCode", "OrderNumber", "ExpectedDeliveryDate","ExpirationDate", "PackageNumber", "Remarks", "DeliveryNumber",
    "OrderedQuantity", "UnitOfMeasure", "CodeByBuyer", "Name", "StreetAndNumber", "CityName", "PostalCode", "StoreNumber", "UnloadingPoint", "ForecastSourceFile")

    for column_name in excel_columns:
        worksheet.write(0, excel_columns.index(column_name), column_name)

    row = 1
    temp_list = []
    buyer_item_list = []
    newest_tmp = []
    position = 1
    test =[]

    for filename in os.listdir(directory_path):
        if filename.split(".")[-1] == "xml":
            # print(filename, "Przeszeł, poszeł")
            path = str(directory_path).replace("/", "\\") + "\\" + filename
            tree = ET.parse(path)
            root = tree.getroot()
            # print(path, "- Gra gitara!")
        else:
            # print(filename, "siem wyjebał")
            continue
            
        for document in root.findall('Order-Header'):
            DocOrderNumber = document.find('OrderNumber').text
            DocOrderDate = document.find('OrderDate').text
            DocOrderDate = (datetime.datetime.strptime(DocOrderDate, '%Y-%m-%d')).date()
            
    
        for line in root.findall('Order-Lines/Line'):
            try:
                BuyerItemCode = line.find('Line-Item/BuyerItemCode').text
            except:
                BuyerItemCode = ""
            try:
                OrderNumber = line.find('Line-Item/OrderNumber').text
            except:
                OrderNumber = ""
            try:
                ExpectedDeliveryDate = line.find('Line-Item/ExpectedDeliveryDate').text
                ExpectedDeliveryDate = (datetime.datetime.strptime(ExpectedDeliveryDate, '%Y-%m-%d')).date()
            except:
                ExpectedDeliveryDate = ""
            try:
                ExpirationDate = line.find('Line-Item/ExpirationDate').text  
                ExpirationDate = (datetime.datetime.strptime(ExpirationDate, '%Y-%m-%d')).date()          
            except:
                ExpirationDate = ""  
            try:
                PackageNumber = line.find('Line-Item/PackageNumber').text
            except:
                PackageNumber = ""  
            try:
                Remarks = line.find('Line-Item/Remarks').text
            except:
                Remarks = "" 
            try:
                DeliveryNumber = line.find('Line-Item/DeliveryNumber').text
            except :
                DeliveryNumber = ""    
            try:
                OrderedQuantity = line.find('Line-Item/OrderedQuantity').text
                OrderedQuantity = int(float(OrderedQuantity))
            except:
                OrderedQuantity = ""
            try:
                UnitOfMeasure = line.find('Line-Item/UnitOfMeasure').text
                if UnitOfMeasure == 'C62' or 'ST':
                    UnitOfMeasure = 'szt'
            except:
                UnitOfMeasure = ""
            try:
                CodeByBuyer = line.find('Line-Parties/DeliveryPoint/CodeByBuyer').text
            except:
                CodeByBuyer = ""
            try:
                Name = line.find('Line-Parties/DeliveryPoint/Name').text
            except:
                Name = ""
            try:
                StreetAndNumber = line.find('Line-Parties/DeliveryPoint/StreetAndNumber').text
            except:
                StreetAndNumber = ""
            try:
                CityName = line.find('Line-Parties/DeliveryPoint/CityName').text
            except:
                CityName = ""
            try:
                PostalCode = (line.find('Line-Parties/DeliveryPoint/PostalCode').text).replace(' ', '-')
                
            except:
                PostalCode = ""
            try:
                StoreNumber = line.find('Line-Parties/DeliveryPoint/StoreNumber').text
            except:
                StoreNumber = ""
            try:    
                UnloadingPoint = line.find('Line-Parties/DeliveryPoint/UnloadingPoint').text
            except:
                UnloadingPoint = ""
            
            line_to_write = [position, DocOrderNumber, DocOrderDate, BuyerItemCode, OrderNumber, ExpectedDeliveryDate, ExpirationDate, PackageNumber, Remarks, DeliveryNumber, OrderedQuantity, UnitOfMeasure, 
            CodeByBuyer, Name, StreetAndNumber, CityName, PostalCode, StoreNumber, UnloadingPoint, filename]    
                
            temp_list.append(line_to_write)
            
            position += 1
    #     print("Position po pliku:", filename, ":", position) 
    # print("Position całkowita:", position)     
    temp_list.sort(key = lambda OrderDate: OrderDate[2])

    for key in temp_list:
        if not key[3] in buyer_item_list:
            buyer_item_list.append(key[3])
    print("Buyer Item Codes:", buyer_item_list)

    for detal in buyer_item_list:
        for record in temp_list:
            if record[3] == detal:
                newest_tmp.append(record)
                # dl = len(newest_tmp)
                # print("newest_tmp wpis:", record, "aktualna długość:", dl, "\n")
        newest_tmp.sort(key = lambda OrderDate: OrderDate[2], reverse=True)
        newest_document_date = newest_tmp[0][2]
        
        # print("\nnewest_document_date:", newest_document_date)
        # print("Detal:", detal, "\nIlość wpisów tymczasowych:", dl)
        # for iteration in temp_list:
        #     for part in iteration:
        #             worksheet.write(row, iteration.index(part), str(part))
        #     row += 1
        for article in newest_tmp:
            if article[2] == newest_document_date:
                test.append(article) 
                for part in range(0, len(article), 1):
                    if isinstance(article[part], datetime.date):
                        worksheet.write_datetime(row, part, article[part], date_format)
                    elif isinstance(article[part], int):
                        worksheet.write_number(row, part, article[part], int_format)
                    else:
                        worksheet.write_string(row, part, str(article[part]), string_format)
                row += 1
        # print("newest_tmp:", newest_tmp, "dla detalu:", detal, "\n")
        newest_tmp.clear()
    # for wpis in test:
    #     print("ID wpisu:", test.index(wpis), wpis, "\n")
    print("\n\nPodsumowanie wpisów do excela:")
    for d in buyer_item_list:
        ile = 0
        for z in test:
            if d == z[3]:
                ile = ile + 1
        print("tego:", d, "jest, tyle:", ile)    
    print("Ilośc wpisów do excela:", len(test))
        
    workbook.close()
#metoda wysyłająca e-mail wraz z ładowaniem załączników
def send(key):
    recipients_list = []
    client_list = []
    for client in key[1:]:
        client_list.append(client.strip())
    print("KLIENCI W GRUPIE:", client_list)
    print("GRUPA:", key[0].strip())
    for rec_group in rec_groups:
        if rec_group[0] == key[0].strip():
            recipients_list_pre = str(rec_group[1:]).strip()[2:-2].split(",")
            for exception in recipients_list_pre:
                recipients_list.append(exception.strip())
            print("ADRESACI W GRUPIE:", recipients_list)

    fromaddr = "prognozy@xzxzxzxzxzx.pl"
    password = "ZxSpectrum83K"
    today = date.today()       
    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients_list)
    msg['Subject'] = str("Plik wygenerowany: " + str(today) + " dla klienta: " + str(client_list))
    
    body = str("\n\n\n\n\n\nJest to wiadomość generowana automatycznie.\nProszę na nią nie odpowiadać.")
    msg.attach(MIMEText(body, 'plain'))

    ct = 0
    aoa = len(os.listdir(direction_path_c))
    try:
        for filename in os.listdir(direction_path_c):
            for client in client_list:
                if filename.lstrip("SL_").rstrip(".xslx") == client:
                    attachment = open(direction_path_c + filename, "rb")
                    name = basename(filename)
                    part = MIMEBase('application', 'xlsx')
                    part.set_payload((attachment).read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % name)
                    msg.attach(part)
                    ct += 1
                    print("Załącznik:", ct, "/", aoa, filename, "- został załadowany.")          
    except:
        print("Brak załączników")
      
    text = msg.as_string()
    server = smtplib.SMTP('smtp.electropoli.pl', 587)
    server.login(fromaddr, password)
    server.sendmail(fromaddr, recipients_list, text)
    server.quit()
    print("Wiadomość została wysłana!")
   
for directory in picked_directories_c:
    parser(directory)

for it in range(0, param_GA, 1):
    send(keys[it])
