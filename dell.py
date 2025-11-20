import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from multiprocessing import Pool, cpu_count
import openpyxl
import time
import re

def dedupe_model_name(name):
    words = name.split()
    half = len(words)//2
    if len(words) >= 2 and words[:half] == words[half:half*2]:
        return " ".join(words[:half])
    return name

def get_dell_data(service_tag):
    options = uc.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-service-autorun")
    options.add_argument("--password-store=basic")
    driver = uc.Chrome(options=options, version_main=0)

    wait = WebDriverWait(driver, 25)

    try:
        driver.get("https://www.dell.com/support/home/pl-pl")

        try:
            cookie_btn = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
            cookie_btn.click()
        except:
            pass

        search = wait.until(EC.presence_of_element_located((By.ID, "mh-search-input")))
        search.send_keys(service_tag)
        search.send_keys(Keys.ENTER)

        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "h1")))

        spec_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "review-specs-drawer-trigger"))
        )
        spec_btn.click()
        time.sleep(2) 

        soup = BeautifulSoup(driver.page_source, "html.parser")

        raw_model = soup.find("h1").get_text(" ", strip=True) if soup.find("h1") else None
        data = {
            "Serial": service_tag,
            "Model": dedupe_model_name(raw_model) if raw_model else None,
            "CPU": None,
            "RAM": None,
            "Dysk": None,
            "Gwarancja": None,
            "ZlyST": False
        }

        # --- CPU / RAM / DYSK ---
        items = soup.find_all("button", class_="dds__accordion__button")
        for item in items:
            text = item.get_text(" ", strip=True)

            # CPU: ignoruj wpisy takie jak w liscie, bo moga byc karty graficzn/wifi intela itd, mozna dodawac do listy po przecinku 
            if (("intel" in text.lower() or "ryzen" in text.lower())
                    and "label" not in text.lower()
                    and not any(k in text.lower() for k in ["wi-fi", "wireless", "bluetooth", "card", "network", "la bel", "graphics", "graficzny", "mod", "klawiatura", "technologia", "etykieta", "graficzna"])
                    and not data["CPU"]):
                data["CPU"] = text

            # RAM
            if "gb" in text.lower() and "mhz" in text.lower() and not data["RAM"]:
                data["RAM"] = text

            # Dysk
            disk_keywords = ["ssd", "nvme", "hdd", "solid state drive", "hard drive"]
            if any(k.lower() in text.lower() for k in disk_keywords) and not data["Dysk"]:
                data["Dysk"] = text

        #Gwarancja
        try:
            warr_div = wait.until(EC.presence_of_element_located((By.ID, "tt_warstatus_text")))
            data["Gwarancja"] = warr_div.get_attribute("innerText").strip()
        except:
            # Jeśli brak elementu gwarancji > zły Service Tag prawdopodobnie
            data["ZlyST"] = True

        return data

    finally:
        driver.quit()


def save_excel(data, file="dell_output.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Model", "Serial", "Specyfikacja", "Gwarancja"])

    for row in data:
        if row.get("ZlyST"):
            spec = "zly ST mordo (chyba)"
        else:
            parts = []
            if row.get("CPU"):
                parts.append(row["CPU"])
            if row.get("RAM"):
                parts.append(row["RAM"])
            if row.get("Dysk"):
                parts.append(row["Dysk"])
            spec = "; ".join(parts)
        ws.append([row.get("Model"), row.get("Serial"), spec, row.get("Gwarancja")])

    wb.save(file)
    print(f"📁 Dane zapisane do: {file}")


if __name__ == "__main__":
    service_tags = ["ST"]
    results = []

    for tag in service_tags:
        print(f"➡ Pobieram dane dla: {tag}")
        try:
            results.append(get_dell_data(tag))
        except Exception as e:
            print(f"❌ Błąd dla {tag}: {e}")
            results.append({
                "Serial": tag,
                "Model": None,
                "CPU": None,
                "RAM": None,
                "Dysk": None,
                "Gwarancja": None,
                "ZlyST": True
            })

    save_excel(results)


