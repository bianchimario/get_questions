import os
import time
import argparse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


def setup_driver():
    """Configura e restituisce un driver Chrome con le opzioni appropriate."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Esegui in background
    chrome_options.add_argument("--window-size=1920,1080")  # Dimensione finestra
    chrome_options.add_argument("--disable-notifications")  # Disabilita notifiche
    chrome_options.add_argument("--disable-infobars")  # Disabilita barre informative
    chrome_options.add_argument("--disable-extensions")  # Disabilita estensioni

    # Inizializza il driver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def capture_element_screenshot(driver, url, selector, output_path, wait_time=10):
    """
    Cattura uno screenshot di un elemento specifico della pagina.
    
    Args:
        driver: WebDriver di Selenium
        url: URL della pagina da cui catturare lo screenshot
        selector: Selettore CSS dell'elemento da catturare
        output_path: Percorso in cui salvare lo screenshot
        wait_time: Tempo massimo di attesa per il caricamento dell'elemento
    
    Returns:
        bool: True se lo screenshot è stato catturato con successo, False altrimenti
    """
    try:
        # Carica la pagina
        driver.get(url)
        
        # Attendi che l'elemento sia visibile
        element = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
        )
        
        # Scorri fino all'elemento
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        
        # Aggiungi un po' di margine sopra l'elemento per non avere lo screenshot troppo attaccato al bordo
        driver.execute_script("window.scrollBy(0, -50);")
        
        # Piccola pausa per assicurarsi che eventuali animazioni siano terminate
        time.sleep(1)
        
        # Cattura lo screenshot dell'elemento
        element.screenshot(output_path)
        
        return True
    
    except TimeoutException:
        print(f"Timeout waiting for element on page: {url}")
        return False
    
    except Exception as e:
        print(f"Error capturing screenshot from {url}: {str(e)}")
        return False


def process_links(links_file, output_dir, selector, topic, start_number=1):
    """
    Processa una lista di link e cattura screenshot per ciascuno.
    
    Args:
        links_file: File con la lista di URL (uno per riga)
        output_dir: Directory in cui salvare gli screenshot
        selector: Selettore CSS dell'elemento da catturare
        topic: Numero del topic per organizzare le immagini
        start_number: Numero da cui iniziare la numerazione degli screenshot
    """
    # Assicurati che la directory di output esista
    topic_dir = os.path.join(output_dir, f"Topic{topic}")
    os.makedirs(topic_dir, exist_ok=True)
    
    # Leggi i link
    if links_file.endswith('.xlsx'):
        # Se è un file Excel, assumiamo che ci sia una colonna 'Link'
        df = pd.read_excel(links_file)
        links = df['Link'].dropna().tolist()
    else:
        # Altrimenti leggi un semplice file di testo con un URL per riga
        with open(links_file, 'r') as f:
            links = [line.strip() for line in f if line.strip()]
    
    # Inizializza il driver
    driver = setup_driver()
    
    try:
        # Processa ogni link
        for i, url in enumerate(links, start=start_number):
            print(f"Processing {i}/{len(links)+start_number-1}: {url}")
            
            output_path = os.path.join(topic_dir, f"{i}.png")
            success = capture_element_screenshot(driver, url, selector, output_path)
            
            if success:
                print(f"  Screenshot saved to {output_path}")
            else:
                print(f"  Failed to capture screenshot")
    
    finally:
        # Chiudi il driver
        driver.quit()


def main():
    # Definisci i parametri da riga di comando
    parser = argparse.ArgumentParser(description='Capture screenshots of elements from a list of URLs')
    parser.add_argument('links_file', help='File with list of URLs')
    parser.add_argument('output_dir', help='Directory to save screenshots')
    parser.add_argument('--selector', default='.question-body', help='CSS selector for the element to capture')
    parser.add_argument('--topic', type=int, default=1, help='Topic number for organizing images')
    parser.add_argument('--start', type=int, default=1, help='Starting number for screenshots')
    
    args = parser.parse_args()
    
    process_links(args.links_file, args.output_dir, args.selector, args.topic, args.start)


if __name__ == "__main__":
    main()