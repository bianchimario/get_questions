# Organizzatore di Screenshot per Domande di Certificazione
# Questo script Python:
# 1. Legge links da un file Excel
# 2. Crea la struttura di cartelle richiesta
# 3. Cattura screenshot dell'elemento specificato con XPath
# 4. Salva gli screenshot con la numerazione corretta

import os
import json
import time
import re
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Configurazione
config = {
    "excel_file": r"C:\Users\mari.bianchi\OneDrive - Reply\Documenti\Tool_Certificazioni\Tool_Certificazioni_streamlit\data\DP-700\database.xlsx",  # Percorso del file Excel
    "base_dir": r"C:\Users\mari.bianchi\Downloads\test_data_8k",             # Cartella base in cui organizzare tutto
    "delay_between_screenshots": 2,  # Secondi di attesa tra screenshot
}

def main():
    try:
        # 1. Leggi i dati dal file Excel
        print(f"Leggendo il file {config['excel_file']}...")
        data = read_excel_file(config['excel_file'])
        
        # 2. Estrai informazioni sul corso e crea la struttura di cartelle
        course_structure = organize_data_by_course(data)
        create_folder_structure(course_structure)
        
        # 3. Cattura gli screenshot per ogni link
        capture_screenshots(course_structure)
        
        print('Operazione completata con successo!')
        
    except Exception as e:
        print(f'Si è verificato un errore: {e}')

def read_excel_file(file_path):
    """Legge il file Excel e restituisce i dati come DataFrame"""
    return pd.read_excel(file_path)

def organize_data_by_course(data):
    """Organizza i dati per corso/certificazione"""
    course_structure = {}
    
    for _, row in data.iterrows():
        if pd.isna(row.get('Link')):
            continue
        
        link = row['Link']
        
        # Estrai il nome del corso (DP-700) dal link
        course_match = re.search(r'exam-([\w-]+)-topic', link)
        course_name = course_match.group(1).upper() if course_match else 'UNKNOWN'
        
        # Estrai il numero del topic dal link o dalla riga
        topic_number = row.get('Topic') if pd.notna(row.get('Topic')) else extract_topic_from_link(link)
        # Converti il topic_number in intero per evitare nomi di cartelle come "Topic1.0"
        topic_number = int(topic_number)
        topic_name = f"Topic{topic_number}"
        
        # Assicurati che le strutture esistano
        if course_name not in course_structure:
            course_structure[course_name] = {
                'topics': {},
                'excel_data': []
            }
        
        if topic_name not in course_structure[course_name]['topics']:
            course_structure[course_name]['topics'][topic_name] = []
        
        # Aggiungi l'informazione della domanda al topic
        course_structure[course_name]['topics'][topic_name].append({
            'numero': int(row['Numero']),
            'link': link
        })
        
        # Archivia anche i dati Excel completi per eventuali usi futuri
        course_structure[course_name]['excel_data'].append(row.to_dict())
    
    return course_structure

def extract_topic_from_link(link):
    """Estrae il numero del topic dal link"""
    match = re.search(r'topic-(\d+)-question', link)
    return int(match.group(1)) if match else 0

def create_folder_structure(course_structure):
    """Crea la struttura di cartelle"""
    print('Creazione della struttura di cartelle...')
    
    # Crea la directory base se non esiste
    os.makedirs(config['base_dir'], exist_ok=True)
    
    # Per ogni corso
    for course_name, course_data in course_structure.items():
        course_path = os.path.join(config['base_dir'], course_name)
        domande_path = os.path.join(course_path, 'Domande')
        
        # Crea la cartella del corso se non esiste
        os.makedirs(course_path, exist_ok=True)
        
        # Copia il file Excel nella cartella del corso
        shutil.copy2(config['excel_file'], os.path.join(course_path, 'database.xlsx'))
        
        # Crea la cartella Domande
        os.makedirs(domande_path, exist_ok=True)
        
        # Crea cartelle per ogni Topic
        for topic_name in course_data['topics']:
            topic_path = os.path.join(domande_path, topic_name)
            os.makedirs(topic_path, exist_ok=True)
    
    print('Struttura di cartelle creata con successo!')

def capture_screenshots(course_structure):
    """Cattura gli screenshot"""
    print('Avvio della cattura degli screenshot...')
    
    # Configura Chrome - rimuoviamo la definizione fissa della dimensione della finestra
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Impostiamo una dimensione più ampia per evitare problemi di visualizzazione
    #chrome_options.add_argument("--window-size=3840,2160")  # Risoluzione 4K
    chrome_options.add_argument("--window-size=7680,4320") # Risoluzione 8K
    
    # Inizializza il browser
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    try:
        # Per ogni corso
        for course_name, course_data in course_structure.items():
            print(f"Elaborazione del corso: {course_name}")
            
            # Per ogni topic
            for topic_name, questions in course_data['topics'].items():
                print(f"Elaborazione di {course_name} - {topic_name}")
                
                # Per ogni domanda nel topic
                for question in questions:
                    numero = question['numero']
                    link = question['link']
                    screenshot_path = os.path.join(
                        config['base_dir'],
                        course_name,
                        'Domande',
                        topic_name,
                        f"{numero}.png"
                    )
                    
                    # Controlla se lo screenshot esiste già
                    if os.path.exists(screenshot_path):
                        print(f"Screenshot per {course_name} {topic_name} Domanda {numero} già esistente.")
                        continue
                    
                    print(f"Cattura screenshot per {course_name} {topic_name} Domanda {numero}...")
                    
                    try:
                        # Vai all'URL
                        driver.get(link)
                        
                        # Attendi che l'elemento con la classe specificata sia visibile
                        wait = WebDriverWait(driver, 30)
                        element = wait.until(EC.visibility_of_element_located(
                            (By.XPATH, '//*[contains(concat( " ", @class, " " ), concat( " ", "discussion-header-container", " " ))]')
                        ))
                        
                        # Cattura lo screenshot dell'elemento completamente
                        # Scrolliamo fino all'elemento ma con un offset verso l'alto per evitare che la navbar lo copra
                        driver.execute_script("arguments[0].scrollIntoView(true);", element)
                        # Aggiungiamo un ulteriore scroll verso l'alto per evitare che la navbar fissa copra l'elemento
                        driver.execute_script("window.scrollBy(0, -100);")
                        
                        # Assicuriamoci che l'intero elemento sia visibile aspettando un attimo
                        time.sleep(0.5)
                        
                        # Catturiamo le dimensioni reali dell'elemento
                        size = element.size
                        
                        # Cattura lo screenshot dell'elemento
                        element.screenshot(screenshot_path)
                        print(f"Screenshot salvato in: {screenshot_path} con dimensioni {size['width']}x{size['height']}px")
                        
                    except Exception as e:
                        print(f"Errore durante la cattura dello screenshot per {course_name} {topic_name} Domanda {numero}: {e}")
                    
                    # Attendi un po' per evitare richieste troppo frequenti
                    time.sleep(config['delay_between_screenshots'])
    
    finally:
        # Chiudi il browser
        driver.quit()
    
    print('Tutti gli screenshot sono stati catturati!')

if __name__ == "__main__":
    main()