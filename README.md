#  Automated CRM Lead Extractor & Router

An automated background service built with **Python** to extract lead data from a dynamic web portal and route it directly into a CRM (Bitrix24) via API. This project demonstrates web scraping resilience, state management, and API integration.

##  Architecture & Workflow

1. **Web Scraping (Selenium):** Bypasses dynamic DOM elements to extract new lead data, handling timeouts and network instability with robust retry mechanisms.
2. **Data Processing (Pandas):** Cleans and structures the extracted information (names, phones, emails, vehicle models).
3. **Round-Robin Distribution:** Implements a local state-tracking system (JSON) to distribute leads evenly among sales representatives based on their physical location and online status.
4. **CRM Integration (REST API):** Pushes the processed leads into the CRM, creating Contact and Deal entities while handling rate limits (HTTP 429/500) using exponential backoff.
5. **Anti-Duplication Cache:** Maintains a 24-hour local cache to prevent processing the same lead twice.

##  Key Features

- **Headless Browser Automation:** Engineered to run efficiently in the background without memory leaks (`kill_orphan_chrome_processes`).
- **Resilient API Calls:** Utilizes `urllib3.util.retry` for exponential backoff during API instability.
- **Environment Security:** Sensitive credentials and webhook URLs are strictly managed via `.env`.
- **Asynchronous Queuing:** Leads captured outside of business hours are saved to a local queue (`pending_leads.json`) and processed automatically the next morning.

##  Tech Stack

- **Language:** Python 3
- **Automation & Scraping:** Selenium WebDriver, WebDriver Manager
- **Data Manipulation:** Pandas
- **API Communication:** Requests, urllib3
- **Configuration:** python-dotenv

##  How to Run

1. Clone this repository:
   git clone https://github.com/luiz103817/honda-lead-bot.git
   cd honda-lead-bot
  
2. Install the required dependencies:
   pip install pandas selenium webdriver-manager python-dotenv requests


3. Configure your environment variables:
   Create a .env file in the root directory based on the .env.example:
    
    CRM_WEBHOOK_URL=https://your-crm-instance.com/rest/userid/token/
    CRM_DRIVE_FOLDER_ID=12345
    PORTAL_USER=your_email@domain.com
    PORTAL_PASS=your_secure_password


4. Execute the main script:
   
   python main.py

