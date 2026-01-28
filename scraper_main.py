import datetime
from playwright.sync_api import sync_playwright
from playwright_stealth import stealth_sync
from dataclasses import dataclass, field
import csv 
import pandas as pd
import lmstudio as lm
import os
from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential
from io import StringIO
import json
import requests
import win32com.client as win32

@dataclass
class Ausschreibung:
    text: str
    link: str


def scrape_google(query):
    with sync_playwright() as p:
        
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = browser.new_page()

        stealth_sync(page)
        
        # Navigate to Google
        page.goto(f"https://www.google.com/search?q={query}")
        
        # Wait for results to load
        page.wait_for_selector('[data-ved]')
        
        # Extract search results
        results = page.query_selector_all('h3')
        links = page.query_selector_all('a[jsname="UWckNb"]')
        #print(links)
        
        for result in results:
            print(result.inner_text())
        for link in links:
            print(link.get_attribute('href'))
        
        browser.close()

def scrape_google_multiple_pages(query):
    all_results = []
    all_links = []
    with sync_playwright() as p:
        
        browser = p.chromium.launch(headless=False)
        
        page = browser.new_page(user_agent="'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36') Chrome/85.0.4183.121 Safari/537.36'")
        #stealth_sync(page)
        page.set_default_timeout(0)
        counter = 0
        while(True):
        # Navigate to Google
         page.goto(f"https://www.google.com/search?q={query}&start={counter}")
        
        # Wait for results to load
         page.wait_for_selector('[data-ved]')
        
        # Extract search results
         results = page.query_selector_all('h3')
         links = page.query_selector_all('a[jsname="UWckNb"]')
        
         for result in results:
             #print(result.inner_text())
             all_results.append(result.inner_text())
         for link in links:
            #print(link.get_attribute('href'))
            all_links.append(link.get_attribute('href'))
         counter = counter + 10
        
         page_result = page.evaluate('() => document.documentElement.innerText') 
         if "Es wurden keine mit deiner Suchanfrage -" in page_result:
          break
         
        browser.close()

        result_dict = dict(zip(all_results,all_links))
        return result_dict

def write_to_csv(dict : dict, filename):
    with open(filename+".csv", 'w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')  # Explicitly set comma delimiter
        
        # Write header
        writer.writerow(['Ergebnis', 'Link'])
        
        # Write data
        for key, value in dict.items():
            writer.writerow([key, value])

def direct_write(csv_content, filename):
    with open(filename+".csv", 'w', newline='', encoding='utf-8-sig') as file:
        file.write(csv_content)


def ai_filter(data):
    client = AzureOpenAI(api_version="2024-12-01-preview",    azure_endpoint="https://adam-m74ujfie-eastus2.openai.azure.com/",   api_key="DUTTkFnYfFoTFifWtPPJ3iaSe642rSVjxzZ3OZQSUwOCqknwc3ooJQQJ99BBACHYHv6XJ3w3AAAAACOGnOGG")

    
    prompt = "Die Daten nach meinen Sätzen sind Google Suchergebnisse und Links für eine Suche. In dieser Suche habe ich nach Ausschreibungen im \
        SAP Finance Bereich gesucht. Bitte filtere alle Ergebnisse raus, die nichts mit SAP oder mit Ausschreibungen zu tun haben. Deine \
            Antwort soll bitte nur das gefilterte Ergebnis sein, welches man so direkt wieder in eine csv Datei schreiben kann. Column 1 soll Ergebnis, Column 2 soll Link heißen.\
                Nutze ; als delimiter!"
    response = client.chat.completions.create(
            model="gpt-5-mini", 
            messages=[
                {"role": "user", "content": prompt + " "+ str(data)}
            ]
        )
    return response.choices[0].message.content
    
def ai_summarize_link(link):

    page_content = ""
    with sync_playwright() as p:
        
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(user_agent="'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36') Chrome/85.0.4183.121 Safari/537.36'")
        page.set_default_timeout(5000)
        filename = ""
        if ".pdf" in link:
            try:
                with page.expect_download() as download_info:
                    page.goto(link)
                    download = download_info.value
                    path = download.path()
                    filename = download.suggested_filename
                    download.save_as(filename)

            except:
                try:
                    page.goto(link)
                    page.wait_for_selector('embed')
                    embed = page.get_by_role("embed").__str__()
                    #print(embed)

                    firstIndex = embed.find("='")
                    secondIndex = embed.find('>')
                    pdf_link = embed[firstIndex+2:secondIndex-1]
                    #print(link)

                    response = requests.get(pdf_link)
                    file = open("file.pdf", "wb")
                    file.write(response.content)
                    file.close()
                    filename = "file.pdf"
                except:
                    return "Page could not be reached"

        page_content = ""
        if filename != "":
            if filename != "file.pdf":
                file = open(filename,"r", encoding="utf-8-sig")
                page_content = file.read()
                file.close()
            else:
                file = open(filename, "rb")
                page_content = file.read()
                file.close()
                
        else:
            counter = 0
        # Navigate to Page
            try:
                page.goto(link)
            except:
                return "Page could not be reached"
        
            page_content = page.query_selector("body").inner_html()
    


    try:
        
        client = AzureOpenAI(api_version="2024-12-01-preview",    azure_endpoint="https://adam-m74ujfie-eastus2.openai.azure.com/",   api_key="DUTTkFnYfFoTFifWtPPJ3iaSe642rSVjxzZ3OZQSUwOCqknwc3ooJQQJ99BBACHYHv6XJ3w3AAAAACOGnOGG")

        if filename == "":
            prompt = "Hierauf folgt der HTML Content einer Webseite über eine Ausschreibung. Bitte fasse den Inhalt zusammen mit allen wichtigen \
        Infos, z.B. Kosten, Technologien usw. Bitte keine Floskeln am Anfang und am Ende, nur die Zusammenfassung. Bitte achte darauf alles \
            zu untersuchen. Das Ergebnis wird in einer csv-Datei landen als Feld, achte auf die Formatierung. Keine Sonderzeichen außer Bindestrich. \
                Nutze diese Formatierung: \
        [Auftraggeber] - [Projekttitel]: [Kurze Beschreibung]. Status: [Vergabestatus]. Gewinner: [Firma oder \"Offen\"]. Wert: [Betrag oder \"Nicht bekannt\"]. Ort: [Stadt]. Besonderheiten: [Falls relevant] \
            Nutze nicht mehr als 400 Zeichen. Keine Zeilenwechsel oder Anführungszeichen. \
                Falls die Seite, die du zusammenfasst, keine Infos über die Ausschreibung enthält, dann lasse die Felder, die nicht vorhanden sind leer. \
                    Außerdem kennzeichne dies in der Kurzbeschreibung. Falls es einen Status gibt: Packe diesen Status zusätzlich an den Anfang der Beschreibung und markiere \
                        ihn mit zwei Sternchen, also so: **XYZ** , wobei XYZ der Status ist. Falls es einen Wert gibt, dann mache das Gleiche wie beim Status \
                            aber nutze zwei Dollarzeichen, also $$123$$ , wobei 123 der Wert ist! Entferne Ausschreibungen, die nicht aus diesem Jahr sind."
        else:
            prompt = "Hierauf folgt der Inhalt einer PDF über eine Ausschreibung. Bitte fasse den Inhalt zusammen mit allen wichtigen \
        Infos, z.B. Kosten, Technologien usw. Bitte keine Floskeln am Anfang und am Ende, nur die Zusammenfassung. Bitte achte darauf alles \
            zu untersuchen. Das Ergebnis wird in einer csv-Datei landen als Feld, achte auf die Formatierung. Keine Sonderzeichen außer Bindestrich. \
                Nutze diese Formatierung: \
        [Auftraggeber] - [Projekttitel]: [Kurze Beschreibung]. Status: [Vergabestatus]. Gewinner: [Firma oder \"Offen\"]. Wert: [Betrag oder \"Nicht bekannt\"]. Ort: [Stadt]. Besonderheiten: [Falls relevant] \
            Nutze nicht mehr als 400 Zeichen. Keine Zeilenwechsel oder Anführungszeichen. \
                Falls die Seite, die du zusammenfasst, keine Infos über die Ausschreibung enthält, dann lasse die Felder, die nicht vorhanden sind leer. \
                    Außerdem kennzeichne dies in der Kurzbeschreibung Falls es einen Status gibt: Packe diesen Status (also ist die Ausschreibung frei oder vergeben) \
                        zusätzlich an den Anfang der Beschreibung und markiere \
                        ihn mit zwei Sternchen, also so: **XYZ** , wobei XYZ der Status ist. Falls es einen Wert (wie viel wird gezahlt) gibt, dann mache das Gleiche wie beim Status \
                            aber nutze zwei Dollarzeichen, also $$123$$ , wobei 123 der Wert ist! Entferne Ausschreibungen, die nicht aus diesem Jahr sind."
        response = client.chat.completions.create(
            model="gpt-5-mini", 
            messages=[
                {"role": "user", "content": prompt + " "+ str(page_content)}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(e)
        return "AI could not summarize, most likely token count exceeded!"
        

def delta_check(data, link: str):
    
    
    for row in data:
        if row["Link"].strip() == link.strip():
            print("Same")
            return True
    
    return False

def ai_analysis():
    with open("postfilter.csv", 'r', newline='', encoding='utf-8-sig') as file:
        reader = csv.reader(file, delimiter=";")
        file = open("final.csv","r", newline='',encoding='utf-8-sig')
        old_reader = csv.DictReader(file,delimiter=';')
        data = list(old_reader)
        
        temp = ["Search Result", "Link", "Open", "Value", "Summary"]
        with open("new_final.csv","w",newline='',encoding='utf-8-sig') as output:
            writer = csv.writer(output, delimiter=";")
            
            counter = 0
            for row in reader:
                if counter == 0:
                    counter = counter + 1
                    writer.writerow(temp)
                    continue
                temp = row
                if delta_check(data,row[1]):
                    continue
                response = ai_summarize_link(temp[1])
                free = findStringBetween(response, "**")
                money = findStringBetween(response, "$$")
                temp.append(free)
                temp.append(money)
                temp.append(response)
                writer.writerow(temp)
                counter = counter +1


def findStringBetween(response : str, marker : str):
    firstIndex = response.find(marker)

    if firstIndex == -1:
        return "No result"
    
    secondIndex = response.find(marker, firstIndex+1)

    if secondIndex == -1:
        return "No result"
    
    substring = response[firstIndex+2:secondIndex]
    return substring



def send_mail():
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')

    #outlook = win32.Dispatch('outlook.application')

    file = open("mails.txt","r")
    recipients = file.readline().strip()
    file.close()

    currentDate = datetime.datetime.now()

    mail = outlook.CreateItem(0)
    mail.To = recipients


    mail.Subject = 'Ausschreibungen für '+currentDate.strftime("%d/%m/%Y")
    mail.Body = 'Ausschreibungen'
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    cwd = os.getcwd()
    attachment  = cwd + "\\final.csv"
    mail.Attachments.Add(attachment)

    mail.Send()
    
def refresh_final():
    new = open("new_final.csv","r",newline='',encoding='utf-8-sig')
    new_reader = csv.reader(new,delimiter=";")
    old = open("final.csv","w",newline='',encoding='utf-8-sig')
    old_writer = csv.writer(old,delimiter=";")
    
    for row in new_reader:
        old_writer.writerow(row)
        
    new.close()
    old.close()
    os.remove("new_final.csv")
    

def main():
    result_dict = scrape_google_multiple_pages("Ausschreibung \"SAP\" datasphere OR SAC OR BW/4HANA OR SAP OR S_4HANA OR BW OR BI OR S/4HANA OR HANA -Blog -Karriere -Jobs -Consulting -Agentur -Anbieter -Whitepaper -Download -Testversion -Produktseite -Werbung -Pressemitteilung -Veranstaltung -Demo -Webinar -site:sap.com -site:help.sap.com -site:linkedin.com -site:youtube.com -site:indeed.com -site:glassdoor.com -site:xing.com -site:kununu.com -site:heise.de -site_computerwoche.de")
    print("Scraped Google")
    write_to_csv(result_dict, "prefilter")
    print("Prefilter CSV written")
    filtered_result = ai_filter(result_dict)
    print("AI filtered")
    #print(filtered_result)
    direct_write(filtered_result,"postfilter")
    print("Postfilter CSV written")
    ai_analysis()
    print("AI Finished")
    refresh_final()
    print("Completed Delta")
    send_mail()
    print("Mail sent")
    print("Done")
    



if __name__ == "__main__":
    main()