
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time


# Load Excel file
df = pd.read_excel("combined_funds.xlsx")  # Replace with your filename

# Setup ChromeDriver (make sure chromedriver is in PATH)
driver = webdriver.Chrome()

for idx, row in df.iterrows():
    fund_code = str(row[0])  # Assuming first column is the fund code
    url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fund_code}"
    
    try:
        driver.get(url)
        time.sleep(2)  # Wait for JS content to load; can be optimized with WebDriverWait
        
        # Find the "Kategorisi" span
        elements = driver.find_elements(By.XPATH, "//li[contains(., 'Kategorisi')]/span")
        category = elements[0].text.strip() if elements else "Not found"
        # Assign to third column (index 2)
        df.iat[idx, 2] = category
        
        # Extract Yatırımcı Sayısı (Kişi)
        investor_element = driver.find_elements(By.XPATH, "//li[contains(., 'Yatırımcı Sayısı (Kişi)')]/span")
        investor_count = investor_element[0].text.strip() if investor_element else "Not found"
        df.iat[idx, 12] = investor_count  # Column 12: Yatırımcı Sayısı
        # print(f"{fund_code}: {category}")

        # Extract Pazar Payi
        marketshare_element = driver.find_elements(By.XPATH, "//li[contains(., 'Pazar Payı')]/span")
        marketshare = marketshare_element[0].text.strip() if marketshare_element else "Not found"
        df.iat[idx, 13] = marketshare  # Column 12: Yatırımcı Sayısı
        # print(f"{fund_code}: {category}")

        # Fonun Risk Değeri
        risk_el = driver.find_elements(By.XPATH, "//td[contains(text(), 'Fonun Risk Değeri')]/following-sibling::td")
        risk_value = risk_el[0].text.strip() if risk_el else "Not found"
        df.iat[idx, 14] = risk_value

        # Fonun Risk Değeri
        status_element = driver.find_elements(By.XPATH, "//td[contains(text(), 'Platform İşlem Durumu')]/following-sibling::td")
        fund_status = status_element[0].text.strip() if status_element else "Not found"
        df.iat[idx, 15] = fund_status

        print(f"{fund_code} | Kategori: {category} | Yatırımcı: {investor_count} | Pazar Payi: {marketshare} | Risk: {risk_value} | Status: {fund_status} ")


    except Exception as e:
        print(f"Error with fund {fund_code}: {e}")
        df.iat[idx, 2] = "Error"
        df.iat[idx, 12] = "Error"
        df.iat[idx, 13] = "Error"
        df.iat[idx, 14] = "Error"
        df.iat[idx, 15] = "Error"

driver.quit()

# Save the updated file
df.to_excel("combined_funds_1.xlsx", index=False)














# import pandas as pd
# import os

# # Get all Excel files in the current directory
# excel_files = [f for f in os.listdir() if f.endswith('.xlsx') or f.endswith('.xls')]

# # Initialize an empty list to hold dataframes
# df_list = []

# # Read and collect all Excel files into the list
# for file in excel_files:
#     df = pd.read_excel(file)
#     df_list.append(df)

# # Concatenate all dataframes into one
# combined_df = pd.concat(df_list, ignore_index=True)

# # Save to a new Excel file
# combined_df.to_excel('combined_funds.xlsx', index=False)

# print("✅ All Excel files merged into 'combined_funds.xlsx'")