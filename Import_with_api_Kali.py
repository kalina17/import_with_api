# -*- coding: utf-8 -*-
"""
Created on Tue Apr  8 18:36:13 2025

@author: waran
"""


import mailerlite as MailerLite
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
from selenium.webdriver.support.ui import Select
import zipfile
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl
from openpyxl.styles import PatternFill
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import numpy as np


# Loguję się na platformę elearningową, dane do logowania należaóby pewnie
# trzymać gdzies stacjonarnie w pliku, a nie wpisywać w kodzie.
driver = webdriver.Chrome()
driver.get([adres])
login_field = driver.find_element(By.ID, "user_login")
login_field.send_keys(login)
passw = driver.find_element(By.ID, "user_pass")
passw.send_keys(password)

# Sleep jest na okolicznosć ładowania się strony. To chyba nie jest eleganckie,
# ale okazuje się mniej zawodne niż "Until(EC.element_to_be_clickable)"
time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="wp-submit"]').click()

# Czyszczę wybór pól, który mógł być wprowadzony z interfejsu graficznego przez
# innego użytkownika i definiuję pola, które potrzebuję zassać.
kolumny = driver.find_element(By.ID, "columns")
kolumny.clear()
kolumny.send_keys("user_login,user_email,first_name,last_name"
                  ",akcept-regulamin_10_11")
# Generuję dzisiejszą datę, żeby skrypt pobrał rekordy, które pojawiły się do
# dzis i ustawiam datę, od kiedy ma pobierać.  Można tu oczywiscie dać
# przestrzeń na jakis input użytkownika.
current_date = datetime.now().strftime('%d.%m.20%y')
data_startowa = driver.find_element(By.XPATH, '//*[@id="from"]')
data_startowa.send_keys("16.04.2025")
data_koncowa = driver.find_element(By.XPATH, '//*[@id="to"]')
data_koncowa.send_keys(current_date)

# Na koniec skrypt klika przycisk "pobierz", a potem musi odczekać, bo to się
# kilka minut generuje
driver.find_element(By.XPATH, '//*[@id="acui_download_csv_wrapper"]'
                    '/td/input[1]').click()
time.sleep(600)

# Ponieważ to jest plik, który się sciąga do folderu pobrane, trzeba go
# znaleźć.  Robię to po początku nazwy, ponieważ każdy eksport dostaję inny
# kod, który znajduje sie w końcówce nazwy pliku.  Jako download_dir
# trzeba podać scieżkę do swojego folderu, do którego domyslnie
# sciągają się pliki z przeglądarki.
file_prefix = "user-export"
files = os.listdir(download_dir)

# Znajduję ostatnio pobrany plik i proceduję tylko, jesli został pobrany maks
# kwadrans temu. Żeby przypadkiem nie przesłały się jakies archiwalne dane.
def file_modified_recently(file_path):
    modification_time = os.path.getmtime(file_path)
    last_modified = datetime.fromtimestamp(modification_time)
    current_time = datetime.now()
    time_difference = current_time - last_modified
    return time_difference < timedelta(minutes=900)


# Znajduję najnowszy plik i go wczytuję
matching_files = [file for file in files if file.startswith(file_prefix)]

if matching_files:
    most_recent_file = max(matching_files, key=lambda x: os.path.getmtime(os.path.join(download_dir, x)))

# Konstruuję pełną scieżkę niezbędną do wczytania pliku i go wczytuję
    file_path = os.path.join(download_dir, most_recent_file)
    df = pd.read_csv(file_path)

# Lista na odpowiedzi ze strony API
responses = [[]

# Hasło do API, pewnie najbardziej poprawnie byłoby je trzymać
# w postaci jakiego pliku stacjonarnie na komputerze, a nie podawać w kodzie
# ja po prostu tego nie wrzucałem do żadnego repozytorium, więc mi to
# nie robiło.
client = MailerLite.Client({
  'api_key': [Apikey]
})

# Przesyłam dane do mailerlite.  Ponieważ wszystkie dodawane osoby pochodzą
# z rejestracji na platformie e-learnignowej, osoby są dodawane z kodem grupy
# "użytkownicy e-learning", to jest to pole "groups".  Oprócz tego osoby różnią
# się tym, czy wyraziły zgodę na przesyłanie newslettera (informacja o tym jest
# w kolumnie 'akcept-regulamin_10_11').  Jesli ta kolumna jest pusta, to znaczy
# że nie wyraziły zgody i stąd różnica w przesyłanych danych
# Można się tu pokusić o wysłanie całego batcha, ale w moim przypadku przy
# 50 do 100 adresach wysyłanych na raz, nie było to potrzebne.
for index, row in df.iterrows():
    email = row['user_email']
    imie = row['first_name']
    nazwisko = row['last_name']

    try:
        if not pd.isnull(row['akcept-regulamin_10_11']):
            response = client.subscribers.create(
                email,
                fields={"name": imie, "last_name": nazwisko, "newsletter": "x"},
                groups=[132344751384954388],
                status="active",
                ip_address='1.2.3.4',
                optin_ip='1.2.3.4'
            )
        else:
            response = client.subscribers.create(
                email,
                fields={"name": imie, "last_name": nazwisko},
                groups=[132344751384954388],
                status="active",
                ip_address='1.2.3.4',
                optin_ip='1.2.3.4'
            )
# Dodaję status odpowiedzi z API, żeby móc przesledzic, jesli pojawiły się
# jakies błędy.
        responses.append({
            "email": email,
            "status_code": response.status_code if hasattr(response, 'status_code') else 200,
            "success": True,
            "response": str(response)
        })

    except Exception as e:
        responses.append({
            "email": email,
            "status_code": getattr(e, 'status_code', 'N/A'),
            "success": False,
            "error": str(e)
        })
    
