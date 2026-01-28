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
import time
import requests


with sync_playwright() as p:
        
        browser = p.chromium.launch(headless=False)
        
        page = browser.new_page(user_agent="'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36') Chrome/85.0.4183.121 Safari/537.36'")
        #stealth_sync(page)
        page.set_default_timeout(0)
        counter = 0
        while(True):
        # Navigate to Google
         page.goto(f"https://www.stadtwerke-erfurt.de/pb/site/swe/get/documents_E616006357/swe/documents/Konzern/Ausschreibungen/2021/Ver%C3%B6ffentlichung%20S001-2021%20Transformation%20SAP%20S4%20HANA.pdf")
         page.wait_for_selector('embed')
         embed = page.get_by_role("embed").__str__()
         print(embed)

         firstIndex = embed.find("='")
         secondIndex = embed.find('>')
         link = embed[firstIndex+2:secondIndex-1]
         print(link)

         response = requests.get(link)
         file = open("test.pdf", "wb")
         file.write(response.content)
         file.close()
         #print(embed[0].get_properties())
         #pdf_link = embed.get_attribute("original-url")
         #print(pdf_link)
         time.sleep(20)
    
