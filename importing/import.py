import os
import sys
import io
import time
import pyautogui
import streamlit as st
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# ✅ Fix encoding issue (Windows CMD compatibility)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

class Start:
    def __init__(self):
        try:
            # Get today's date in 'YYYY-MM-DD' format
            self.today_date = datetime.now().strftime("%Y-%m-%d")

            # File path to upload
            self.file_path = f"\\\\192.168.15.241\\admin\\ACTIVE\\jlborromeo\\CBS HOME LOAN\\FOR VOLARE\\FOR UPLOAD\\CBS HOMELOAN AS {self.today_date}.xlsx"

            # Check if file exists
            if not os.path.exists(self.file_path):
                self.show_message("File Not Found", f"The file '{self.file_path}' does not exist.\nThe process will close in 5 seconds.")
                sys.exit(1)

            # Initialize Chrome WebDriver
            self.setup_browser()

            # Perform Login
            self.login()

            # Navigate to Import Manager
            self.navigate_to_import_manager()

            # Upload File
            self.upload_file()

            # Wait for completion notification
            self.wait_for_import_completion()

            print("✅ Import process completed successfully.")

        except Exception as e:
            print(f"❌ Error occurred: {e}")
        finally:
            # Ensure the browser closes
            if hasattr(self, 'driver'):
                self.driver.quit()

    def setup_browser(self):
        """ Initializes Chrome browser with necessary options """
        print("🚀 Launching browser...")  # ✅ Safe with UTF-8 encoding
        options = Options()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.driver = webdriver.Chrome(options=options)
        self.driver.maximize_window()

    def login(self):
        """ Handles the login process """
        print("🔑 Logging in...")
        self.driver.get("https://texxen-voliappe3.spmadridph.com/admin/")
        
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "inputName"))).send_keys("JMBORROMEO")
        self.driver.find_element(By.ID, "inputPassword").send_keys("$PMadr!d1234")
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))).click()

        # Wait until login is successful
        WebDriverWait(self.driver, 10).until(EC.url_changes("https://texxen-voliappe3.spmadridph.com/admin/"))
        print("✅ Login successful.")

    def navigate_to_import_manager(self):
        """ Navigates to the Import Manager page """
        print("📂 Navigating to Import Manager...")
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[span[text()=' Import Manager ']]"))).click()
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-path='ImportManager.Batch']"))).click()
        print("✅ Import Manager opened.")

    def upload_file(self):
        """ Uploads the specified file """
        print(f"📤 Uploading file: {self.file_path}")

        file_input = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        file_input.send_keys(self.file_path)
        pyautogui.press("enter")

        # Click "Upload & Proceed"
        WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btnSubmit']"))).click()

        # Click "Proceed to Import"
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='filePreviewSubmit']"))).click()

        print("✅ File uploaded and processing started.")

    def wait_until_element_appears(self, xpath, check_interval=5):
        """
        Waits indefinitely until the specified element appears on the page.
        
        Args:
            xpath: XPath of the target element.
            check_interval: Time (seconds) between each check.

        Returns:
            WebElement: The found element.
        """
        print(f"⏳ Waiting for element: {xpath}")
        while True:
            try:
                element = self.driver.find_element(By.XPATH, xpath)
                if element.is_displayed():
                    print(f"✅ Element appeared: {xpath}")
                    return element  # Successfully found element
            except NoSuchElementException:
                pass  # Element not found, continue waiting

            print(f"🔄 Still waiting... Checking again in {check_interval} seconds.")
            time.sleep(check_interval)  # Wait before checking again
            time.sleep(70)
            
    def wait_for_import_completion(self):
        """ Waits indefinitely for the import process to complete """
        print("⏳ Waiting for import completion...")

        # Step 1: Wait for the bulk notification container (if it appears)
        try:
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@id, 'bulk-notifier-completed-content')]"))
            )
            print("✅ Notification container detected.")
        except TimeoutException:
            print("⚠ Notification container did not appear within 60 seconds, but will continue waiting.")

        # Step 2: Wait indefinitely for the exact success message
        complete_message_xpath = f"//label[text()='Completed sending import batch CBS HOMELOAN AS {self.today_date}.xlsx']"

        self.wait_until_element_appears(complete_message_xpath)
        print("✅ Import process completed successfully.")
        time.sleep(50)

    def show_message(self, title, message):
        """ Displays a pop-up message using tkinter """
        root = tk.Tk()
        root.withdraw()  # Hide root window
        messagebox.showerror(title, message)

# Run the automation
if __name__ == "__main__":
    Start()
