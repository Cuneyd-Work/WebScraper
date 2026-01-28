import csv
from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential
from playwright.sync_api import sync_playwright
from playwright_stealth import stealth_sync
import requests

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
                            aber nutze zwei Dollarzeichen, also $$123$$ , wobei 123 der Wert ist!"
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
                            aber nutze zwei Dollarzeichen, also $$123$$ , wobei 123 der Wert ist!"
        response = client.chat.completions.create(
            model="o4-mini", 
            messages=[
                {"role": "user", "content": prompt + " "+ str(page_content)}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(e)
        return "AI could not summarize, most likely token count exceeded!"
        


def ai_analysis():
    with open("postfilter.csv", 'r', newline='', encoding='utf-8-sig') as file:
        reader = csv.reader(file, delimiter=";")
        
        temp = ["Search Result", "Link", "Open", "Value", "Summary"]
        with open("final.csv","w",newline='',encoding='utf-8-sig') as output:
            writer = csv.writer(output, delimiter=";")
            
            counter = 0
            for row in reader:
                if counter == 0:
                    counter = counter + 1
                    writer.writerow(temp)
                    continue
                temp = row
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