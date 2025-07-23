
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
import time
import random
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration
WAIT_TIME = 10
RESTART_DRIVER_EVERY = 100
MAX_RETRIES = 3
INPUT_FILE = "combined_funds.xlsx"
OUTPUT_FILE = "combined_funds_1.xlsx"

def create_driver():
    """Create a new Chrome driver with optimized options."""
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # Commented out for debugging
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-plugins')
    # options.add_argument('--disable-images')  # Commented out for debugging
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def safe_find_element(driver, xpath, wait_time=WAIT_TIME):
    """Safely find an element with proper error handling."""
    try:
        element = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        return element.text.strip() if element else "Not found"
    except (TimeoutException, NoSuchElementException) as e:
        logger.debug(f"Element not found with xpath: {xpath} - {e}")
        return "Not found"
    except Exception as e:
        logger.warning(f"Unexpected error finding element: {e}")
        return "Error"

def debug_page_content(driver, fund_code):
    """Debug function to inspect page content."""
    try:
        # Wait for page to load
        time.sleep(3)
        
        # Check if page loaded correctly
        page_title = driver.title
        current_url = driver.current_url
        logger.info(f"Page title: {page_title}")
        logger.info(f"Current URL: {current_url}")
        
        # Try to find any li elements
        li_elements = driver.find_elements(By.TAG_NAME, "li")
        logger.info(f"Found {len(li_elements)} li elements")
        
        # Try to find any span elements
        span_elements = driver.find_elements(By.TAG_NAME, "span")
        logger.info(f"Found {len(span_elements)} span elements")
        
        # Try to find any td elements
        td_elements = driver.find_elements(By.TAG_NAME, "td")
        logger.info(f"Found {len(td_elements)} td elements")
        
        # Look for specific text content
        page_source = driver.page_source
        if "Kategorisi" in page_source:
            logger.info("✓ Found 'Kategorisi' in page source")
        else:
            logger.warning("✗ 'Kategorisi' not found in page source")
            
        if "Yatırımcı Sayısı" in page_source:
            logger.info("✓ Found 'Yatırımcı Sayısı' in page source")
        else:
            logger.warning("✗ 'Yatırımcı Sayısı' not found in page source")
            
        # Try alternative XPath patterns
        alt_patterns = [
            "//span[contains(text(), 'Kategorisi')]",
            "//td[contains(text(), 'Kategorisi')]",
            "//div[contains(text(), 'Kategorisi')]",
            "//*[contains(text(), 'Kategorisi')]"
        ]
        
        for pattern in alt_patterns:
            try:
                elements = driver.find_elements(By.XPATH, pattern)
                if elements:
                    logger.info(f"✓ Found elements with pattern: {pattern}")
                    break
            except Exception as e:
                logger.debug(f"Pattern failed: {pattern} - {e}")
                
    except Exception as e:
        logger.error(f"Error in debug function: {e}")

def clear_cookies_and_cache(driver):
    """Clear cookies and cache to avoid tracking issues."""
    try:
        driver.delete_all_cookies()
        driver.execute_script("window.localStorage.clear();")
        driver.execute_script("window.sessionStorage.clear();")
    except Exception as e:
        logger.warning(f"Could not clear cookies/cache: {e}")

def scrape_fund_data(driver, fund_code, url, retry_count=0):
    """Scrape data for a single fund with retry logic."""
    try:
        # Check if browser is alive
        try:
            driver.current_url
        except WebDriverException:
            logger.warning(f"Driver crashed for fund {fund_code}. Need to restart.")
            raise WebDriverException("Driver crashed")

        logger.info(f"Navigating to: {url}")
        driver.get(url)
        
        # Clear cookies periodically
        if retry_count == 0:  # Only on first attempt
            clear_cookies_and_cache(driver)

        # Wait for page to load
        time.sleep(3)
        
        # Debug the first few funds to understand the page structure
        if retry_count == 0:  # Only debug on first attempt
            debug_page_content(driver, fund_code)

        # Extract all required data with multiple xpath attempts
        data = {}
        
        # Try multiple XPath patterns for each element
        category_xpaths = [
            "//li[contains(., 'Kategorisi')]/span",
            "//span[contains(text(), 'Kategorisi')]/following-sibling::span",
            "//td[contains(text(), 'Kategorisi')]/following-sibling::td",
            "//*[contains(text(), 'Kategorisi')]/following-sibling::*"
        ]
        
        investor_xpaths = [
            "//li[contains(., 'Yatırımcı Sayısı (Kişi)')]/span",
            "//span[contains(text(), 'Yatırımcı Sayısı')]/following-sibling::span",
            "//td[contains(text(), 'Yatırımcı Sayısı')]/following-sibling::td",
            "//*[contains(text(), 'Yatırımcı Sayısı')]/following-sibling::*"
        ]
        
        market_share_xpaths = [
            "//li[contains(., 'Pazar Payı')]/span",
            "//span[contains(text(), 'Pazar Payı')]/following-sibling::span",
            "//td[contains(text(), 'Pazar Payı')]/following-sibling::td",
            "//*[contains(text(), 'Pazar Payı')]/following-sibling::*"
        ]
        
        risk_xpaths = [
            "//td[contains(text(), 'Fonun Risk Değeri')]/following-sibling::td",
            "//span[contains(text(), 'Risk Değeri')]/following-sibling::span",
            "//*[contains(text(), 'Risk Değeri')]/following-sibling::*"
        ]
        
        status_xpaths = [
            "//td[contains(text(), 'Platform İşlem Durumu')]/following-sibling::td",
            "//span[contains(text(), 'İşlem Durumu')]/following-sibling::span",
            "//*[contains(text(), 'İşlem Durumu')]/following-sibling::*"
        ]
        
        # Try each xpath pattern until one works
        data['category'] = try_multiple_xpaths(driver, category_xpaths)
        data['investor_count'] = try_multiple_xpaths(driver, investor_xpaths)
        data['market_share'] = try_multiple_xpaths(driver, market_share_xpaths)
        data['risk_value'] = try_multiple_xpaths(driver, risk_xpaths)
        data['fund_status'] = try_multiple_xpaths(driver, status_xpaths)
        
        return data
        
    except Exception as e:
        if retry_count < MAX_RETRIES:
            logger.warning(f"Retrying fund {fund_code} (attempt {retry_count + 1}/{MAX_RETRIES}): {e}")
            time.sleep(random.uniform(1, 3))
            return scrape_fund_data(driver, fund_code, url, retry_count + 1)
        else:
            logger.error(f"Failed to scrape fund {fund_code} after {MAX_RETRIES} attempts: {e}")
            return {
                'category': 'Error',
                'investor_count': 'Error',
                'market_share': 'Error',
                'risk_value': 'Error',
                'fund_status': 'Error'
            }

def try_multiple_xpaths(driver, xpaths):
    """Try multiple XPath patterns and return the first successful result."""
    for xpath in xpaths:
        try:
            elements = driver.find_elements(By.XPATH, xpath)
            if elements and elements[0].text.strip():
                result = elements[0].text.strip()
                logger.debug(f"Success with xpath: {xpath} -> {result}")
                return result
        except Exception as e:
            logger.debug(f"Failed xpath: {xpath} - {e}")
            continue
    return "Not found"

def main():
    """Main scraping function."""
    # Load Excel file
    try:
        df = pd.read_excel(INPUT_FILE)
        logger.info(f"Loaded {len(df)} funds from {INPUT_FILE}")
        
        # Fix pandas dtype issues by converting columns to object type
        df.iloc[:, 2] = df.iloc[:, 2].astype('object')   # Category
        df.iloc[:, 11] = df.iloc[:, 11].astype('object') # Investor Count
        df.iloc[:, 12] = df.iloc[:, 12].astype('object') # Market Share
        df.iloc[:, 13] = df.iloc[:, 13].astype('object') # Risk Value
        df.iloc[:, 14] = df.iloc[:, 14].astype('object') # Fund Status
        
    except Exception as e:
        logger.error(f"Could not load Excel file: {e}")
        return

    driver = create_driver()
    successful_scrapes = 0
    
    try:
        # Test with first few funds only for debugging
        test_limit = 55555  # Remove this after debugging
        
        for idx, row in df.iterrows():
            if idx >= test_limit:  # Remove this after debugging
                break
                
            fund_code = str(row.iloc[0])
            url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fund_code}"

            # Restart browser every N items
            if idx > 0 and idx % RESTART_DRIVER_EVERY == 0:
                driver.quit()
                driver = create_driver()
                logger.info(f"Restarted driver at index {idx}")

            # Scrape fund data
            data = scrape_fund_data(driver, fund_code, url)
            
            # Assign to dataframe columns
            df.iat[idx, 2] = data['category']          # Column 3: Category
            df.iat[idx, 11] = data['investor_count']   # Column 12: Investor Count
            df.iat[idx, 12] = data['market_share']     # Column 13: Market Share
            df.iat[idx, 13] = data['risk_value']       # Column 14: Risk Value
            df.iat[idx, 14] = data['fund_status']      # Column 15: Fund Status

            # Log progress
            if data['category'] != 'Error':
                successful_scrapes += 1
            
            logger.info(f"[{idx+1}/{len(df)}] {fund_code} | "
                       f"Category: {data['category']} | "
                       f"Investors: {data['investor_count']} | "
                       f"Market Share: {data['market_share']} | "
                       f"Risk: {data['risk_value']} | "
                       f"Status: {data['fund_status']}")

            # Add random sleep to reduce server load
            time.sleep(random.uniform(0.5, 2.0))

            # Save progress every 50 funds
            if (idx + 1) % 50 == 0:
                df.to_excel(OUTPUT_FILE, index=False)
                logger.info(f"Progress saved at index {idx + 1}")

    except KeyboardInterrupt:
        logger.info("Scraping interrupted by user")
    except Exception as e:
        logger.error(f"Unexpected error during scraping: {e}")
    finally:
        driver.quit()
        logger.info("Browser closed")

    # Save the final updated file
    df.to_excel(OUTPUT_FILE, index=False)
    logger.info(f"Scraping completed! Successfully scraped {successful_scrapes}/{len(df)} funds")
    logger.info(f"Results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
