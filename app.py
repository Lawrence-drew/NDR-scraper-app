import tkinter as tk
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import socket
import pandas as pd
from selenium.webdriver.support.ui import Select



num_pages = 0

def check_internet_connection():
    try:
        # Try to establish a connection to a well-known server (e.g., Google's DNS)
        socket.create_connection(("www.google.com", 80))
        return True  # Internet connection is available
    except OSError:
        return False  # No internet connection


# Define the function to perform web scraping
def scrape_webpage():
        

        if not check_internet_connection():
            # Display an error message on the GUI
            status_label.config(text="No internet connection. Please check your connection and retry.")
            return
        global num_pages
        # Get the webpage URL from the entry field
        url = "https://ndr.phis3project.org.ng/Identity/Account/Login?ReturnUrl=%2F"
        
        # Get login credentials from user
        username = username_entry.get()
        password = password_entry.get()

        # Get the number of pages to scrape from the user input field
        num_pages_str = num_pages_entry.get()
        
        # Validate that num_pages_str is a positive integer
        if not num_pages_str.isdigit() or int(num_pages_str) <= 0:
            status_label.config(text="Please enter a valid positive integer for the number of pages.")
            return
        
        num_pages = int(num_pages_str)
        
        # Get the desired CSV file name from the user
        file_name = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        
        if file_name:
            # Create a Selenium WebDriver instance (you may need to specify the driver path)
            driver = webdriver.Chrome()
            
            # Navigate to the login page
            driver.get(url)
            
            # Fill in the login form
            username_field = driver.find_element(By.XPATH,"//input[@type='email']")  # Replace with actual field name
            password_field = driver.find_element(By.XPATH,"//input[@type='password']")  # Replace with actual field name
            
            username_field.send_keys(username)
            password_field.send_keys(password)
            password_field.send_keys(Keys.RETURN)
            
            # Wait for login to complete (you can adjust the sleep time as needed)
            time.sleep(1)
            
            # Navigate to the page with the table you want to scrape
            driver.get('https://ndr.phis3project.org.ng/uploads/partner-uploads')
            
            # Scraping the table (modify this part according to your scraping script)
            # Example: You can use Pandas to scrape tables if they are in HTML format
            # Replace this example with your actual scraping logic

            all_data = pd.DataFrame(columns = ['IP/Date', 'Batch/Facility', 'Status', 'Total', 'Fails', 'Passes', 'Pending','Details'])
            sleep_between_pages = 6


            select = Select(driver.find_element(By.XPATH,"//select"))
            select.select_by_index(2)

            for page_number in range(1, num_pages + 1):
                # Scroll down to trigger dynamic loading
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                # Wait for the dynamic content to load (adjust timeout as needed)
                wait = WebDriverWait(driver, 2)
                wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr")))

                table = driver.find_elements(By.XPATH, "//table//tbody/tr")  # Replace with actual XPath
                for row in table:
                    data = []
                    data_rows = []
                    cells = row.find_elements(By.XPATH, "td")
                    row_data = [cell.text for cell in cells]

                    data.append(row_data)
                    data_rows.extend(data)

                    page_data = pd.DataFrame(data_rows, columns=['IP/Date', 'Batch/Facility', 'Status', 'Total', 'Fails', 'Passes', 'Pending','Details'])
                    all_data = pd.concat([all_data, page_data], ignore_index=True)

            
                # df = pd.DataFrame(data)
                if page_number < num_pages:
                    # Check if there's a next page button and click it if it exists
                    next_page_button = driver.find_element(By.ID, "uploadDataTable_next")
                    if next_page_button and next_page_button.is_enabled():
                        next_page_button.click()
                        # You can add a delay here to avoid overloading the website
                        time.sleep(sleep_between_pages)
                        print(f"Scraping page {page_number}")
                    else:
                        break 
     

            
            # Save the scraped data to an Excel file
            all_data.to_excel(file_name, index=False)
            status_label.config(text=f"Data saved to {file_name}")

        else:
            status_label.config(text="File not saved")

        # Close the WebDriver
        driver.quit()
        
# Create the GUI window
window = tk.Tk()
window.title("Web Scraping GUI")

# Create and pack GUI elements
username_label = tk.Label(window, text="Username:")
username_label.pack()

username_entry = tk.Entry(window, width=40)
username_entry.pack()

password_label = tk.Label(window, text="Password:")
password_label.pack()

password_entry = tk.Entry(window, width=40, show="*")  # Hide password characters
password_entry.pack()

# Add a label and entry for specifying the number of pages to scrape
num_pages_label = tk.Label(window, text="Number of Pages to Scrape:")
num_pages_label.pack()

num_pages_entry = tk.Entry(window, width=5)
num_pages_entry.pack()

scrape_button = tk.Button(window, text="Scrape and Save to Excel", command=scrape_webpage)
scrape_button.pack()

status_label = tk.Label(window, text="")
status_label.pack()

# Start the GUI main loop
window.mainloop()
