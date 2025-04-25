# -*- coding: utf-8 -*-
"""
Last version 25.04.2025

The source of the data is a form created with Elementor (a WordPress plugin), and the target application is a mailing tool – MailerLite.

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


# Logowanie na platformę elearningową, dane do logowania należałoby docelowo trzymać w oddzielnym pliku
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

# Czyszcznie wyboru pól, który mógłby być wprowadzony z interfejsu graficznego przez
# innego użytkownika i definiowanie pola, które potrzeba zassać.
kolumny = driver.find_element(By.ID, "columns")
kolumny.clear()
kolumny.send_keys("user_login,user_email,first_name,last_name"
                  ",akcept-regulamin_10_11")
# Generowanie dzisiejszej daty, żeby skrypt pobrał rekordy, które pojawiły się do
# dzis i ustawienie daty, od kiedy ma pobierać. Może docelowo również przestrzeń na input użytkownika?
current_date = datetime.now().strftime('%d.%m.20%y')
data_startowa = driver.find_element(By.XPATH, '//*[@id="from"]')
data_startowa.send_keys("16.04.2025")
data_koncowa = driver.find_element(By.XPATH, '//*[@id="to"]')
data_koncowa.send_keys(current_date)

# Na koniec skrypt klika przycisk "pobierz", a potem musi odczekać, bo to się
# kilka minut generuje i stąd sleep.
driver.find_element(By.XPATH, '//*[@id="acui_download_csv_wrapper"]'
                    '/td/input[1]').click()
time.sleep(600)

# Ponieważ to jest plik, który się sciąga do folderu pobrane, trzeba go
# znaleźć.  Najłatwiej znaleźć go po początku nazwy, ponieważ każdy eksport dostaje inny
# kod, który znajduje sie w końcówce nazwy pliku. Jako download_dir
# trzeba podać scieżkę do swojego folderu, do którego domyslnie
# sciągają się pliki z przeglądarki.
file_prefix = "user-export"
files = os.listdir(download_dir)

# Należy znaleźć ostatnio pobrany plik i procedować tylko, jesli został pobrany maks
# kwadrans temu - aby przypadkiem nie przesłały się jakies archiwalne dane.
def file_modified_recently(file_path):
    modification_time = os.path.getmtime(file_path)
    last_modified = datetime.fromtimestamp(modification_time)
    current_time = datetime.now()
    time_difference = current_time - last_modified
    return time_difference < timedelta(minutes=900)


# Znajdowanie najnowszego pliku i wczytanie go:
matching_files = [file for file in files if file.startswith(file_prefix)]

if matching_files:
    most_recent_file = max(matching_files, key=lambda x: os.path.getmtime(os.path.join(download_dir, x)))

# Tworzenie pełnej scieżki niezbędnej do wczytania pliku, oraz wykonanie wczytania:
    file_path = os.path.join(download_dir, most_recent_file)
    df = pd.read_csv(file_path)

# Lista na odpowiedzi ze strony API
responses = [[]

# Hasło do API, pewnie najbardziej poprawnie byłoby je trzymać
# w postaci jakiego pliku stacjonarnie na komputerze, a nie podawać w kodzie - docelowo do poprawy

client = MailerLite.Client({
  'api_key': [Apikey]
})

# Przesłąnie danych do mailerlite. Wszystkie dodawane osoby pochodzą
# z rejestracji na platformie e-learnignowej, osoby są dodawane z kodem grupy
# "użytkownicy e-learning", to jest to pole "groups".  Oprócz tego osoby różnią
# się tym, czy wyraziły zgodę na przesyłanie newslettera (informacja o tym jest
# w kolumnie 'akcept-regulamin_10_11').  Jesli ta kolumna jest pusta, to znaczy
# że nie wyraziły zgody i stąd różnica w przesyłanych danych.
# Można się tu pokusić o wysłanie całego batcha, ale na razie w przypadku
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
# Dodanie statusu odpowiedzi z API, żeby móc prześledzic, jesli pojawiły się błędy:
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
    
