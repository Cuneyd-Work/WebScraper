import win32com.client as win32
import datetime
import os
import csv

def delta_check(data, link: str):
    
    
    for row in data:
        if row["Link"].strip() == link.strip():
            print("Same")
            return True

with open("test_postfilter.csv", 'r', newline='', encoding='utf-8-sig') as file:
        new_reader = csv.reader(file, delimiter=";")
        new_dict_reader = csv.DictReader(file, delimiter=";")
        file = open("test_final.csv","r", newline='',encoding='utf-8-sig')
        old_reader = csv.DictReader(file,delimiter=';')
        data = list(old_reader)
        
        temp = ["Search Result", "Link", "Open", "Value", "Summary"]
        with open("test_new_final.csv","w",newline='',encoding='utf-8-sig') as output:
            writer = csv.writer(output, delimiter=";")
            
            counter = 0
            for row in new_reader:
                if counter == 0:
                    counter = counter + 1
                    writer.writerow(temp)
                    continue
                temp = row
                print(row[1])
                if delta_check(data,row[1]):
                    continue
                response = "Test"
                
                
                temp.append(response)
                writer.writerow(temp)
                counter = counter +1
