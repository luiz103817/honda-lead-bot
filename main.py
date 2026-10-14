import re
import requests
import pandas as pd
import os
import sys
import time
import json
import subprocess 
from datetime import datetime, timedelta
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv
import logging

# set up paths
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

LOGS_DIR = os.path.join(SCRIPT_DIR, 'sys_data')
if not os.path.exists(LOGS_DIR):
    os.makedirs(LOGS_DIR)

LOG_FILE = os.path.join(LOGS_DIR, 'runtime_cache.log')

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def custom_print(*args, **kwargs):
    msg = " ".join(map(str, args))
    if "[ERROR" in msg.upper() or "[ERRO" in msg.upper():
        logging.error(msg)
    elif "[WARN" in msg.upper() or "[AVISO" in msg.upper():
        logging.warning(msg)
    else:
        logging.info(msg)

def log_uncaught_exceptions(ex_cls, ex, tb):
    logging.critical("Fatal execution failure", exc_info=(ex_cls, ex, tb))
sys.excepthook = log_uncaught_exceptions

# local storage files
SELLERS_FILE = os.path.join(SCRIPT_DIR, 'sellers.xlsx')
DISTRIBUTION_STATE_FILE = os.path.join(SCRIPT_DIR, 'distribution_state.json')
MONTHLY_COUNT_FILE = os.path.join(SCRIPT_DIR, 'leads_count.json')
PROCESSED_CUSTOMERS_FILE = os.path.join(SCRIPT_DIR, 'processed_customers.xlsx')
PENDING_LEADS_FILE = os.path.join(SCRIPT_DIR, 'pending_leads.json')
CACHE_FILE = os.path.join(SCRIPT_DIR, 'history_24h.json')
ENV_FILE = os.path.join(SCRIPT_DIR, '.env')

load_dotenv(dotenv_path=ENV_FILE)

# crm config
DEFAULT_ASSIGNEE_ID = 61622
CRM_WEBHOOK_URL = os.getenv("CRM_WEBHOOK_URL")
CRM_DRIVE_FOLDER_ID = os.getenv("CRM_DRIVE_FOLDER_ID")

if not CRM_DRIVE_FOLDER_ID:
    sys.exit()
else:
    CRM_DRIVE_FOLDER_ID = int(CRM_DRIVE_FOLDER_ID)

# static mappings
CITY_TO_LIST_ID = {
    "Araras": 45, "Itapira": 47, "Leme": 49, "Mogi Guacu": 51, 
    "Mogi Mirim": 53, "Santa Barbara": 55, "Campinas": 57, 
    "Hortolandia": 59, "Sumare": 61, "Nova Odessa": 105, 
    "Pinhal": 180, "Conchal": 182
}

CITY_CODE_MAP = {
    "1019570": "Araras", "1617361": "Leme", "1696878": "Mogi Mirim",
    "1696882": "Itapira", "1696922": "Mogi Guacu", "1708795": "Campinas",
    "1709487": "Sumare", "1748942": "Hortolandia", "1748941": "Santa Barbara"
}

def safe_extract_text(driver, xpath, wait=False):
    # safe text extraction from web elements
    try:
        if wait:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        
        elements = driver.find_elements(By.XPATH, xpath)
        for el in elements:
            text = el.get_attribute("textContent").strip()
            if text:
                return text
        return "Not Informed"
    except (NoSuchElementException, TimeoutException):
        return "Not Informed"

def safe_extract_input_value(driver, xpath, wait=False):
    # safe value extraction for input tags
    try:
        if wait:
            element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
            val = element.get_attribute("value")
        else:
            val = driver.find_element(By.XPATH, xpath).get_attribute("value")
            
        return val.strip() if val else "Not Informed"
    except (NoSuchElementException, TimeoutException):
        return "Not Informed"

def is_lead_recently_processed(unique_id, hours=24):
    # prevents duplicate processing within a 24h window
    if not os.path.exists(CACHE_FILE):
        return False
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            cache = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return False
        
    if unique_id in cache:
        date_str = cache[unique_id]
        try:
            record_date = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
            now = datetime.now()
            if now - record_date < timedelta(hours=hours):
                return True 
        except ValueError:
            pass 
    return False

def log_lead_to_cache(unique_id):
    cache = {}
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass
    cache[unique_id] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, indent=4)
        
def is_working_hours():
    current_hour = datetime.now().hour
    return 8 <= current_hour < 18

def save_pending_lead(client, lead_name):
    # queues leads collected outside working hours
    try:
        with open(PENDING_LEADS_FILE, 'r', encoding='utf-8') as f:
            pending = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pending = []
        
    pending.append({"lead_name": lead_name, "client": client})
    with open(PENDING_LEADS_FILE, 'w', encoding='utf-8') as f:
        json.dump(pending, f, indent=4, ensure_ascii=False)

def process_pending_leads(df_sellers):
    # flushes the pending queue
    if not os.path.exists(PENDING_LEADS_FILE): return
    try:
        with open(PENDING_LEADS_FILE, 'r', encoding='utf-8') as f:
            pending = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError): return
    if not pending: return
    
    remaining = []
    
    for item in pending:
        client = item['client']
        lead_name = item['lead_name']
        assignee_id, assignee_name = get_next_assignee(client['store'], df_sellers)
        contact = find_existing_contact(client["phone"])
        contact_id = contact["ID"] if contact else create_crm_contact(client, assignee_id)
        
        success = False 
        if contact_id:
            deal_id = find_existing_deal(contact_id)
            if deal_id:
                success = True 
            else:
                deal_id = create_crm_deal(client, contact_id, assignee_id)
                if deal_id:
                    success = True 
                    register_lead_count(assignee_name)
                
                try:
                    new_lead_data = {
                        "Processing Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Name": client.get("name", "Not Found"),
                        "Phone": client.get("phone", "Not Found"),
                        "Email": client.get("email", "Not Found"),
                        "Store": client.get("store", "Not Found"),
                        "Deal Type": client.get("deal_type", "Not Found"),
                        "Notes": client.get("notes", "No notes"),
                        "Detailed Model": client.get("detailed_model", "Not Found"),
                        "CPF": client.get("cpf", "Not Informed"),
                        "Color": client.get("color", "Not Informed"),
                        "Assigned_Seller": assignee_name,
                        "CRM Contact ID": contact_id,
                        "CRM Deal ID": deal_id
                    }
                    expected_cols = ["Processing Date", "Name", "Phone", "Email", "Store", "Deal Type", "Notes", "Detailed Model","CPF", "Color", "Assigned_Seller", "CRM Contact ID", "CRM Deal ID"]
                    try:
                        df_processed = pd.read_excel(PROCESSED_CUSTOMERS_FILE)
                    except:
                        df_processed = pd.DataFrame(columns=expected_cols)
                        
                    new_df = pd.DataFrame([new_lead_data], columns=expected_cols)
                    df_processed = pd.concat([df_processed, new_df], ignore_index=True)
                    df_processed.to_excel(PROCESSED_CUSTOMERS_FILE, index=False)
                except Exception: 
                    pass

        if success:
            continue 
        else:
            remaining.append(item)
        
    with open(PENDING_LEADS_FILE, 'w', encoding='utf-8') as f:
        json.dump(remaining, f, indent=4, ensure_ascii=False)

def find_existing_deal(contact_id):
    if not contact_id: return None
    limit_date = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%dT%H:%M:%S")
    params = {
        "filter": {"CONTACT_ID": contact_id, "CATEGORY_ID": 12, ">=DATE_CREATE": limit_date },
        "select": ["ID"]
    }
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}crm.deal.list.json", json=params)
        r.raise_for_status()
        deals = r.json().get("result", [])
        if deals: return deals[0]["ID"] 
    except: pass
    return None 

def kill_orphan_chrome_processes():
    # prevents memory leaks from crashed browser sessions
    try:
        ps_cmd = (
            "Get-WmiObject Win32_Process | "
            "Where-Object { $_.Name -eq 'chrome.exe' -and $_.CommandLine -match 'ChromeProfileRobot' } | "
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
        )
        subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True)
        time.sleep(2)
    except Exception:
        pass

def map_store_by_code(store_code):
    if not store_code: return "Store not found"
    city = CITY_CODE_MAP.get(str(store_code))
    if city == "SBO": return "Santa Barbara"
    return city if city else "Store not found"

def call_crm_api_with_retry(method, params=None, files=None, max_retries=5, initial_delay=5):
    # wrapper with exponential backoff
    url = f"{CRM_WEBHOOK_URL}{method}.json"
    session = requests.Session()
    retry_strategy = Retry(
        total=max_retries, backoff_factor=initial_delay,
        status_forcelist=[429, 500, 502, 503, 504], allowed_methods=["POST", "GET"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    
    for attempt in range(1, max_retries + 1):
        try:
            if files: response = session.post(url, files=files, timeout=60)
            else: response = session.post(url, json=params or {}, timeout=30)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException:
            wait_time = initial_delay * (2 ** (attempt - 1))
            if attempt < max_retries: time.sleep(wait_time)
            else: return None
        finally: session.close()
    return None

def register_lead_count(seller_name):
    current_month = datetime.now().strftime("%Y-%m")
    try:
        with open(MONTHLY_COUNT_FILE, 'r', encoding='utf-8') as f: counts = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError): counts = {}
    
    if current_month not in counts: counts[current_month] = {}
    counts[current_month][seller_name] = counts[current_month].get(seller_name, 0) + 1
    
    with open(MONTHLY_COUNT_FILE, 'w', encoding='utf-8') as f: 
        json.dump(counts, f, indent=2, ensure_ascii=False)

def get_next_assignee(city, df_sellers):
    # round-robin distribution system
    try:
        if os.path.exists(DISTRIBUTION_STATE_FILE):
            with open(DISTRIBUTION_STATE_FILE, 'r', encoding='utf-8') as f: state = json.load(f)
        else: state = {}
    except (FileNotFoundError, json.JSONDecodeError): state = {}

    toggle_key = f"turn_for_{city.lower()}"
    next_type = "LOCAL" if state.get(toggle_key) == "ONLINE" else "ONLINE"
    state[toggle_key] = next_type
    final_seller = None

    if next_type == "LOCAL":
        team = df_sellers[df_sellers['City'].str.lower() == city.lower()].to_dict('records')
        index_key = f"index_local_{city.lower()}"
        if not team: next_type = "ONLINE"
        else:
            idx = state.get(index_key, -1)
            next_idx = (idx + 1) % len(team)
            final_seller = team[next_idx]
            state[index_key] = next_idx

    if next_type == "ONLINE":
        team = df_sellers[df_sellers['City'].str.lower() == 'online'].to_dict('records')
        index_key = "index_online_global" 
        if team:
            idx = state.get(index_key, -1)
            next_idx = (idx + 1) % len(team)
            final_seller = team[next_idx]
            state[index_key] = next_idx

    if not final_seller:
        return DEFAULT_ASSIGNEE_ID, "Default Assignee"

    with open(DISTRIBUTION_STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=4, ensure_ascii=False)

    try: clean_id = str(int(final_seller['CRM_ID']))
    except (ValueError, TypeError):
        clean_id = str(DEFAULT_ASSIGNEE_ID)

    clean_name = str(final_seller['Name']).strip()
    return clean_id, clean_name
    
def normalize_phone(phone): return re.sub(r'\D', '', phone)

def format_phone_with_cc(phone):
    tel = normalize_phone(phone)
    if not tel: return ""
    if len(tel) in [10, 11] and not tel.startswith("55"): return f"55{tel}"
    return tel

def find_existing_contact(phone, email=""):
    if phone and "informed" not in phone.lower():
        params = {"filter": {"PHONE": format_phone_with_cc(phone)},"select": ["ID", "ASSIGNED_BY_ID"]}
        try:
            r = requests.post(f"{CRM_WEBHOOK_URL}crm.contact.list.json", json=params)
            r.raise_for_status()
            contacts = r.json().get("result", [])
            if contacts: return contacts[0]
        except: pass
    
    if email and "informed" not in email.lower() and "@" in email:
        params = {"filter": {"EMAIL": email},"select": ["ID", "ASSIGNED_BY_ID"]}
        try:
            r = requests.post(f"{CRM_WEBHOOK_URL}crm.contact.list.json", json=params)
            r.raise_for_status()
            contacts = r.json().get("result", [])
            if contacts: return contacts[0]
        except: pass

    return None

def create_crm_contact(client, assignee_id):
    name = client.get("name", "").strip()
    phone = client.get("phone", "Not Informed").strip()
    extracted_email = client.get("email", "Not Informed").strip() 

    if not name or "informed" in name.lower(): return None
    
    formatted_phone = format_phone_with_cc(phone) if "informed" not in phone.lower() else ""
    valid_email = extracted_email if "@" in extracted_email and "." in extracted_email else ""

    if not formatted_phone and not valid_email:
        return None

    fields = {
        "NAME": name,
        "ASSIGNED_BY_ID": assignee_id,
        "OPENED": "Y"
    }
    
    if valid_email: fields["EMAIL"] = [{"VALUE": valid_email, "VALUE_TYPE": "WORK"}]
    if formatted_phone: fields["PHONE"] = [{"VALUE": formatted_phone, "VALUE_TYPE": "WORK"}]

    params = {"fields": fields}
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}crm.contact.add.json", json=params)
        r.raise_for_status()
        return r.json().get("result")
    except: return None

def create_crm_deal(client, contact_id, assignee_id):
    city_id = CITY_TO_LIST_ID.get(client['store'])
    custom_fields = {}

    # map custom crm attributes
    if city_id: custom_fields['UF_CRM_1689593052602'] = city_id
    if client.get('deal_type') and 'informed' not in client['deal_type'].lower(): custom_fields['UF_CRM_1703613635957'] = client['deal_type']
    if client.get('detailed_model') and 'informed' not in client['detailed_model'].lower(): custom_fields['UF_CRM_1730900925'] = client['detailed_model']
    if client.get('color') and client['color'] != "Not Informed": custom_fields['UF_CRM_1708101160503'] = client['color']
    
    cpf_val = client.get("cpf")
    if cpf_val and "Informed" not in cpf_val: custom_fields['UF_CRM_1730124165976'] = cpf_val
    
    params = {
        "fields": {
            "TITLE": client["name"], 
            "CONTACT_ID": contact_id, 
            "ASSIGNED_BY_ID": assignee_id, 
            "COMMENTS": client["notes"], 
            "CATEGORY_ID": 12,    
            "STAGE_ID": "C12:NEW", 
            **custom_fields
        }
    }
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}crm.deal.add.json", json=params)
        r.raise_for_status()
        return r.json().get("result")
    except: return None

def search_crm_file(file_name, folder_id):
    params = {"id": folder_id, "filter": {"NAME": file_name}}
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}disk.folder.getchildren.json", json=params)
        r.raise_for_status()
        res = r.json().get("result", [])
        for item in res:
            if item.get("TYPE") == "file" and item.get("NAME") == file_name: return item["ID"], item["NAME"]
    except: pass
    return None, None

def delete_crm_file(file_id, file_name=""):
    params = {"id": file_id}
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}disk.file.delete.json", json=params)
        r.raise_for_status()
        return bool(r.json().get("result"))
    except: return False

def rename_crm_file(file_id, new_name):
    params = {"id": file_id, "newName": new_name}
    try:
        r = requests.post(f"{CRM_WEBHOOK_URL}disk.file.rename.json", json=params)
        r.raise_for_status()
        return bool(r.json().get("result"))
    except: return False

def upload_crm_file(local_path, folder_id, delete_existing=True):
    if not os.path.exists(local_path): return False
    file_name = os.path.basename(local_path)
    
    if delete_existing:
        search_resp = call_crm_api_with_retry("disk.folder.getchildren", {"id": folder_id, "filter": {"NAME": file_name}})
        existing_id = None
        if search_resp and "result" in search_resp:
            for item in search_resp["result"]:
                if item.get("TYPE") == "file" and item.get("NAME") == file_name:
                    existing_id = item["ID"]
                    break
        if existing_id:
            del_resp = call_crm_api_with_retry("disk.file.delete", {"id": existing_id})
            if del_resp and del_resp.get("result"): time.sleep(2)

    upload_url_data = call_crm_api_with_retry("disk.folder.uploadfile", {"id": folder_id, "data": {"NAME": file_name}})
    if not upload_url_data or "result" not in upload_url_data or "uploadUrl" not in upload_url_data["result"]: return False
    upload_url = upload_url_data["result"]["uploadUrl"]

    session = requests.Session()
    retry_strategy = Retry(total=5, backoff_factor=5, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    try:
        with open(local_path, 'rb') as f:
            files = {'file': (file_name, f, 'application/octet-stream')}
            upload_response = session.post(upload_url, files=files, timeout=120)
            upload_response.raise_for_status()
            upload_result = upload_response.json()
            if upload_result and "result" in upload_result and "ID" in upload_result["result"]: return True
    except: pass
    finally: session.close()
    return False

def sync_crm_reports():
    if os.path.exists(PROCESSED_CUSTOMERS_FILE): 
        upload_crm_file(PROCESSED_CUSTOMERS_FILE, CRM_DRIVE_FOLDER_ID)
    
    count_file_name = os.path.basename(MONTHLY_COUNT_FILE)
    current_month = datetime.now().strftime("%Y-%m")
    if not os.path.exists(MONTHLY_COUNT_FILE): return 

    file_id_crm, _ = search_crm_file(count_file_name, CRM_DRIVE_FOLDER_ID)
    if file_id_crm:
        try:
            download_response = requests.post(f"{CRM_WEBHOOK_URL}disk.file.get.json", json={"id": file_id_crm})
            download_response.raise_for_status()
            download_data = download_response.json()
            if "result" in download_data and "DOWNLOAD_URL" in download_data["result"]:
                download_content = requests.get(download_data["result"]["DOWNLOAD_URL"]).text
                try:
                    crm_counts = json.loads(download_content)
                    first_month_crm = next(iter(crm_counts)) if crm_counts else None
                    if first_month_crm and first_month_crm != current_month:
                        archived_name = f"leads_count_{first_month_crm}.json"
                        archived_existing_id, _ = search_crm_file(archived_name, CRM_DRIVE_FOLDER_ID)
                        if archived_existing_id:
                            delete_crm_file(archived_existing_id, archived_name)
                            time.sleep(1)
                        if rename_crm_file(file_id_crm, archived_name):
                            upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID, delete_existing=False)
                        else: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)
                    else: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)
                except: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)
            else: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)
        except: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)
    else: upload_crm_file(MONTHLY_COUNT_FILE, CRM_DRIVE_FOLDER_ID)

# --- Main execution ---
if __name__ == "__main__":
    leads_processed_this_run = 0

    while True:
        kill_orphan_chrome_processes()
        driver = None
        try:
            try:
                df_sellers = pd.read_excel(SELLERS_FILE)
            except FileNotFoundError:
                time.sleep(60)
                continue 

            CHROME_USER_DATA_DIR = r"C:\ChromeProfileRobot" 
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-notifications")
            chrome_options.add_argument("--disable-infobars")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--log-level=3")
            chrome_options.add_argument(f"user-data-dir={CHROME_USER_DATA_DIR}")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

            driver = webdriver.Chrome(service=Service(), options=chrome_options)
            wait = WebDriverWait(driver, 25)

            driver.get("https://myhonda.my.site.com/leads")
            time.sleep(2)
            
            # authentication handling
            if "login" in driver.current_url:
                wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(os.getenv("MYHONDA_USER"))
                driver.find_element(By.ID, "password").send_keys(os.getenv("MYHONDA_PASS"))
                driver.find_element(By.ID, "Login").click()
                wait.until(EC.presence_of_element_located((By.XPATH, "//a[span[text()='LEADBOX']]")))

            leads_list_url = "https://myhonda.my.site.com/myhonda2/s/lead/Lead/00BVP00000FCOFd2AP?Lead-filterId=NewLead"
            MAX_BROWSER_RETRIES = 5
            BROWSER_RETRY_DELAY = 30 
            last_processed_datetime = None
            
            while True:
                try:
                    df_sellers = pd.read_excel(SELLERS_FILE)
                except Exception:
                    pass 

                if is_working_hours():
                    process_pending_leads(df_sellers)

                page_loaded = False
                for attempt in range(1, MAX_BROWSER_RETRIES + 1):
                    try:
                        driver.get(leads_list_url)
                        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//table[@role='grid']/tbody")))
                        time.sleep(5)
                        page_loaded = True
                        break 
                    except (TimeoutException, Exception):
                        if attempt < MAX_BROWSER_RETRIES: time.sleep(BROWSER_RETRY_DELAY)
                
                if not page_loaded:
                    break 
                
                today_str = datetime.now().strftime("%d/%m/%Y")
                leads_to_visit = []
                
                try:
                    table_rows = driver.find_elements(By.XPATH, "//table[@role='grid']/tbody/tr")
                    
                    for i, row in enumerate(table_rows):
                        try:
                            date_element = row.find_element(By.XPATH, ".//span[contains(@class, 'uiOutputDateTime')]")
                            date_text = date_element.get_attribute("textContent").strip()
                            lead_date = date_text.split(" ")[0] 
                            
                            if lead_date != today_str:
                                break
                            
                            lead_datetime = datetime.strptime(date_text, "%d/%m/%Y %H:%M")
                            
                            if last_processed_datetime and lead_datetime < last_processed_datetime:
                                break
                            
                            link_element = row.find_element(By.XPATH, ".//th//a[@data-refid='recordId']")
                            salesforce_id = link_element.get_attribute("data-recordid")
                            table_lead_name = link_element.get_attribute("textContent").strip()
                            lead_url = link_element.get_attribute("href")
                            
                            if is_lead_recently_processed(salesforce_id):
                                continue
                                
                            leads_to_visit.append({
                                "id": salesforce_id,
                                "url": lead_url,
                                "name": table_lead_name,
                                "datetime": lead_datetime 
                            })
                        except Exception:
                            continue 
                except TimeoutException:
                    pass
                    
                if not leads_to_visit:
                    time.sleep(65)
                    continue
                    
                # scraping lead details
                for lead in leads_to_visit:
                    driver.get(lead["url"])
                    time.sleep(4)
                    
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, "//lightning-formatted-name[@slot='primaryField']")))
                    except TimeoutException:
                        continue
                    
                    client = {
                        "unique_id": lead["id"],
                        "name": lead["name"],
                        "notes": "Lead captured from Myhonda bot."
                    }

                    client["email"] = safe_extract_text(driver, "//li[contains(., 'E-mail')]/following-sibling::li")
                    client["phone"] = safe_extract_text(driver, "//li[contains(., 'Celular')]/following-sibling::li")
                    client["cpf"] = safe_extract_text(driver, "//li[contains(., 'CPF/CNPJ')]/following-sibling::li")
                    
                    try:
                        edit_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@name='Lead.EditFlow' or normalize-space()='Editar']"))
                        )
                        driver.execute_script("arguments[0].click();", edit_btn)
                        
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//input[@name='txt_ModelOfInterest']"))
                        )
                        time.sleep(1) 
                        
                        client["detailed_model"] = safe_extract_input_value(driver, "//input[@name='txt_ModelOfInterest']")
                        client["color"] = safe_extract_input_value(driver, "//input[@name='txt_ColorOfInterest']")
                        client["deal_type"] = safe_extract_input_value(driver, "//input[@name='txt_Type']")
                        
                        modal_email = safe_extract_input_value(driver, "//input[@inputmode='email' or @autocomplete='email']")
                        if modal_email != "Not Informed":
                            client["email"] = modal_email
                        client["model"] = client["detailed_model"]
                        
                        close_btn = driver.find_element(By.XPATH, "//button[@title='Cancelar e fechar' or contains(@class, 'slds-modal__close')]")
                        driver.execute_script("arguments[0].click();", close_btn)
                        time.sleep(1.5)
                        
                    except Exception:
                        client["detailed_model"] = "Not Informed"
                        client["model"] = "Not Informed"
                        client["color"] = "Not Informed"
                        client["deal_type"] = "Not Informed"

                    try:
                        details_tab = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Detalhes de Negociação') or @data-label='Detalhes de Negociação']"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", details_tab)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", details_tab)
                        time.sleep(3) 
                    except Exception:
                        pass

                    try:
                        driver.execute_script("""
                            window.scrollTo(0, document.body.scrollHeight);
                            let scrollers = document.querySelectorAll('.slds-scrollable_y, .scroller, .uiScroller');
                            scrollers.forEach(s => s.scrollTop = s.scrollHeight);
                        """)
                        time.sleep(2) 
                    except:
                        pass

                    client["origin"] = safe_extract_text(driver, "//*[contains(text(), 'Origem do lead')]/ancestor::div[contains(@class, 'slds-form-element')]//lightning-formatted-text")
                    
                    store_code = "Not Informed"
                    known_codes = ["1019570", "1617361", "1696878", "1696882", "1696922", "1708795", "1709487", "1748942", "1748941"]
                    
                    try:
                        page_html = driver.page_source
                        page_text = driver.execute_script("return document.body.innerText;") or ""
                        
                        for code in known_codes:
                            if f">{code}<" in page_html or code in page_text:
                                store_code = code
                                break
                    except Exception:
                        pass

                    if store_code != "Not Informed" and store_code.strip() != "":
                        client["store"] = store_code
                    else:
                        client["store"] = "Not Informed"

                    for key, value in client.items():
                        if not value:
                            client[key] = "Not Informed"

                    if client["store"] != "Not Informed":
                        client["store"] = map_store_by_code(client["store"])
                        
                    has_phone = (client["phone"] != "Not Informed" and client["phone"].strip() != "")
                    has_email = (client["email"] != "Not Informed" and "@" in client["email"])
                        
                    if "Not Informed" in client["name"] or "Not Informed" in client["store"] or (not has_phone and not has_email):
                        continue
                        
                    lead_name = client["name"]
                    client_id = lead["id"]
                    
                    if last_processed_datetime is None or lead["datetime"] > last_processed_datetime:
                        last_processed_datetime = lead["datetime"]

                    # advancing internal salesforce flow
                    try:
                        advance_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='AVANÇAR NEGOCIAÇÃO']")))
                        driver.execute_script("arguments[0].click();", advance_btn)
                        time.sleep(2)
                        
                        combobox = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Motivo do Registro' or contains(@id, 'combobox-button')]")))
                        driver.execute_script("arguments[0].click();", combobox)
                        time.sleep(1)
                        
                        other_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//lightning-base-combobox-item[@data-value='Outros']//span[@title='Outros']")))
                        driver.execute_script("arguments[0].click();", other_option)
                        time.sleep(1)
                        
                        register_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Registrar' and contains(@class, 'slds-button_brand')]")))
                        driver.execute_script("arguments[0].click();", register_btn)
                        time.sleep(3) 
                        
                    except Exception:
                        pass

                    # crm integration and reporting
                    if is_working_hours():
                        assignee_id, assignee_name = get_next_assignee(client['store'], df_sellers)
                        contact = find_existing_contact(client["phone"])
                        contact_id = contact["ID"] if contact else create_crm_contact(client, assignee_id)
                        
                        if contact_id:
                            deal_id = find_existing_deal(contact_id)
                            
                            if not deal_id:
                                deal_id = create_crm_deal(client, contact_id, assignee_id)
                                if deal_id:
                                    register_lead_count(assignee_name)
                                    log_lead_to_cache(client_id)
                                try:
                                    new_lead_data = {
                                        "Processing Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "Name": client.get("name", "Not Found"),
                                        "Phone": client.get("phone", "Not Found"),
                                        "Email": client.get("email", "Not Found"),
                                        "Store": client.get("store", "Not Found"),
                                        "Deal Type": client.get("deal_type", "Not Found"),
                                        "Notes": client.get("notes", "No notes"),
                                        "Detailed Model": client.get("detailed_model", "Not Found"),
                                        "CPF": client.get("cpf", "Not Informed"),
                                        "Color": client.get("color", "Not Informed"),
                                        "Assigned_Seller": assignee_name, 
                                        "CRM Contact ID": contact_id,
                                        "CRM Deal ID": deal_id
                                    }
                                    expected_cols = ["Processing Date", "Name", "Phone", "Email", "Store", "Deal Type", "Notes", "Detailed Model","CPF", "Color", "Assigned_Seller", "CRM Contact ID", "CRM Deal ID"]
                                    try:
                                        df_processed = pd.read_excel(PROCESSED_CUSTOMERS_FILE)
                                        for col in expected_cols:
                                            if col not in df_processed.columns: df_processed[col] = pd.NA 
                                    except FileNotFoundError: df_processed = pd.DataFrame(columns=expected_cols)
                                    
                                    new_df = pd.DataFrame([new_lead_data], columns=expected_cols)
                                    df_processed = pd.concat([df_processed, new_df], ignore_index=True)
                                    df_processed.to_excel(PROCESSED_CUSTOMERS_FILE, index=False)
                                except Exception: 
                                    pass
                    else:
                        save_pending_lead(client, lead_name)
                        log_lead_to_cache(client_id)
                    
                    leads_processed_this_run += 1
                    
        except KeyboardInterrupt:
            sys.exit() 
        except Exception:
            time.sleep(30)
        finally:
            if driver:
                try: driver.quit()
                except: pass
            sync_crm_reports()