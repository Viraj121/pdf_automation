import os
import time
import requests  # Import requests for downloading files
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
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,  # Set download directory
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True  # Automatically download PDFs instead of opening them in Chrome
    })
    
    # Use WebDriver Manager to automatically manage ChromeDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    return driver

def wait_for_downloads(download_dir):
    """Wait until all files are downloaded."""
    while True:
        if all([not filename.endswith('.crdownload') for filename in os.listdir(download_dir)]):
            break
        time.sleep(1)  # Wait before checking again

def scroll_into_view(driver, element):
    """Scroll the specified element into view."""
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(1)  # Optional: wait for scrolling to complete

def main():
    # Load form links from Excel
    excel_file = 'data/links.xlsx'
    links = read_excel_links(excel_file)

    # Set up download directory and Selenium WebDriver
    download_dir = os.path.join(os.getcwd(), 'downloads')  # Create a downloads directory
    os.makedirs(download_dir, exist_ok=True)  # Create directory if it doesn't exist
    driver = setup_driver(download_dir)
    
    for link in links:
        print(f"Opening form link: {link}")
        driver.get(link)
        time.sleep(2)  # Wait for page to load
        
        try:
            # Locate the stitch button using XPath and scroll into view before clicking
            stitch_button = driver.find_element(By.XPATH, '//*[@id="stitchpdfqa"]')
            scroll_into_view(driver, stitch_button)  # Scroll to the stitch button
            
            # Get the href attribute from the stitch button (PDF URL)
            pdf_url = stitch_button.get_attribute('href')
            print(f"PDF URL fetched: {pdf_url}")

            if pdf_url:
                # Convert PDF URL to S3 URL if necessary (modify this as per your requirement)
                s3_pdf_url = convert_to_s3_url(pdf_url)  # Implement this function based on your logic
                
                print(f"Converted S3 PDF URL: {s3_pdf_url}")

                # Download the PDF using requests
                response = requests.get(s3_pdf_url)
                
                if response.status_code == 200:
                    pdf_path = os.path.join(download_dir, "downloaded_pdf.pdf")  # Define a path for saving the PDF
                    
                    with open(pdf_path, 'wb') as f:
                        f.write(response.content)  # Write the content of the response to a file
                    
                    print(f"PDF downloaded successfully at: {pdf_path}")

                    # Process the downloaded PDF
                    output_path = os.path.join(download_dir, "output_highlighted_combined.pdf")
                    highlight_year_and_bleed_marks(pdf_path, output_path)
                    print(f"Processed PDF saved at: {output_path}")
                else:
                    print(f"Failed to download PDF. Status code: {response.status_code}")
            else:
                print("No valid PDF URL found.")

        except Exception as e:
            print(f"Error occurred while processing link {link}: {e}")

    driver.quit()

def convert_to_s3_url(pdf_url):
    """Convert the given PDF URL to an S3 URL as needed."""
    # Implement your logic here to convert pdf_url to s3_pdf_url.
    return pdf_url  # Placeholder: replace with actual conversion logic

if __name__ == "__main__":
    main()
