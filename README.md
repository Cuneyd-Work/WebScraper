# WebScraper
Python Webscraper für Ausschreiben


##Benötigte Dateien: 
-scraper_main.py - Hauptprogramm
-mails.txt - Textdatei mit den Mails, an die die resultierende Liste gehen soll; Mails durch Semikolon getrennt

##Benötigte Programme: 
-python3 - Version relativ egal, am Besten die aktuelle stabile
-IDE - Egal welche, ist prinzipiell auch nicht verpflichtend, ist nur leichter es so auszuführen
-Chrome

###Wichtig: Wenn Python installiert wird kann es sein, dass einige Bibliotheken fehlen. Dann wird bei der Ausführung des Programmes ein Fehler angezeigt. Mit pip install die fehlenden Bibliotheken installieren.

##Ausführung: 
Entweder über Kommandozeile oder über IDE das Programm starten. Nachdem das Programm startet öffnet der Chrome-Browser, dort muss höchstwahrscheinlich ein Captcha gelöst werden, dies lösen. Danach scraped das Programm automatisch per Google Suche. Am Ende gibt es die final.csv Datei aus und schickt automatisch eine Mail.

##Wichtige Codestellen:

- api_version, azure_endpoint und api_key Variablen müssen je nach AI Model und Zugang umgeändert werden
- result_dict in main() Methode enthält alle Query Begriffe, je nach Bedürfnis anpassen
- Jede Funktion, die mit AI beginnt, nutzt AI Features. Jede dieser Funktion enthält eine Prompt, diese nicht groß ändern, jedoch können kleine Anpassungen je nach Bedürfnis getätigt werden


###Hinweise:
Wahrscheinlich werden die Ergebnisse mit besseren Models noch passender aussehen.
Außer der scraper_main.py und der mails.txt sind keine der Dateien zur Ausführung notwendig, sie werden während das Programm läuft generiert. Jedoch habe ich sie gepusht,  damit man sich ein Bild machen kann.


####Cüneyd Tas
