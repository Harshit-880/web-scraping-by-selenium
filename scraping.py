import pandas as pd
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from docx import Document

# Set up the WebDriver
options = Options()
options.headless = True  # Run in headless mode (without a browser window)
driver = webdriver.Chrome('D:\task2\chrome-win64')

# Function to scrape data from a URL
def scrape_data(url):
    driver.get(url)
    # Extract data (customize this part to match your needs)
    data = {
        'URL': url,
        # Add other data points you want to scrape (e.g., text, images)
        'Title': driver.title,
        # 'Text': driver.find_element(By.CSS_SELECTOR, 'your-css-selector').text  # Replace with your CSS selector
        'Links': [],
        'Images': []
    }
    # CSS selectors for links and images
    link_selector = 'a'  # CSS selector for anchor elements (links)
    image_selector = 'img'  # CSS selector for image elements
    
    # Extract links
    link_elements = driver.find_elements(By.CSS_SELECTOR, link_selector)
    for link_element in link_elements:
        href = link_element.get_attribute('href')
        if href:
            data['Links'].append(href)
    
    # Extract images
    image_elements = driver.find_elements(By.CSS_SELECTOR, image_selector)
    for image_element in image_elements:
        src = image_element.get_attribute('src')
        if src:
            data['Images'].append(src)
    # Return the data dictionary
    return data

# Function to read URLs from a WordPad (docx) file
def read_urls_from_wordpad(file_path):
    # Use python-docx to read the .docx file
    doc = Document(file_path)
    urls = []
    # Iterate through paragraphs and extract URLs
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if text.startswith('http://') or text.startswith('https://'):
            urls.append(text.strip())
    return urls

# Main script
wordpad_file_path = r'c:\\Users\\garga\\Downloads\\Python Assigment 2.docx'  # Replace with your WordPad file path
urls = read_urls_from_wordpad(wordpad_file_path)

for url in urls:
    print(url)

# List to store scraped data
scraped_data = []

# Scrape data from each URL
for url in urls:
    data = scrape_data(url)
    scraped_data.append(data)

# Convert scraped data to a pandas DataFrame
df = pd.DataFrame(scraped_data)

# Save the DataFrame to an Excel file
df.to_excel('scraped_data.xlsx', index=False)

# Quit the WebDriver
driver.quit()
