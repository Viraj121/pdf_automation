import os
import time
import requests  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains  # For scrolling into view
from excel_reader import read_excel_links
from pdf_processor import highlight_year_and_bleed_marks
from webdriver_manager.chrome import ChromeDriverManager

def setup_driver(download_dir):
    """Set up Selenium WebDriver with options."""
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir, 
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True  
    })
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver


def scroll_into_view(driver, element):
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(1) 


def main():
    excel_file = 'data/links.xlsx'
    links = read_excel_links(excel_file)

    download_dir = os.path.join(os.getcwd(), 'downloads')  

    os.makedirs(download_dir, exist_ok=True)  # Create directory if it doesn't exist
    driver = setup_driver(download_dir)
    
    for link in links:
        print(f"Opening form link: {link}")
        driver.get(link)
        
        try:
            stitch_button = driver.find_element(By.XPATH, '//*[@id="stitchpdfqa"]')
            scroll_into_view(driver, stitch_button) 
            pdf_url = stitch_button.get_attribute('href')
            print("Stitch PDF: ",pdf_url)

            if pdf_url:
                response = requests.get(pdf_url)
                
                if response.status_code == 200:
                    pdf_path = os.path.join(download_dir, "original.pdf") 
                    
                    with open(pdf_path, 'wb') as f:
                        f.write(response.content) 

                    print(f"PDF downloaded {pdf_path}")

                    output_path = os.path.join(download_dir, "highlighted.pdf")
                    highlight_year_and_bleed_marks(pdf_path, output_path, pdf_url)
                    print(f"Processed PDF saved at: {output_path}")

                else:
                    print(f"Failed to download PDF. Status code: {response.status_code}")
            else:
                print("No valid PDF URL found.")

        except Exception as e:
            print(f"Error occurred while processing link {link}: {e}")

    driver.quit()

if __name__ == "__main__":
    main()
