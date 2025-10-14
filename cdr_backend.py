import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    NoSuchElementException,
    WebDriverException,
)


class CdrAutomation:
    def __init__(self, browser_name: str, headless=False):
        self.user_id = "your_user_id"
        self.passcode = "your_password"
        self.search_value = "Mountain Private Equity Partners"
        self.driver_location = f"C:\\Drivers\\{browser_name}"
        self.map_url = "https://example.com/login"  # <-- Replace with your actual URL
        self.headless = headless
        self.browser_name = browser_name
        self.retry_attempts = 3
        self.retry_delay = 2

    # -------------------------------------------------------------------------
    # Browser setup
    # -------------------------------------------------------------------------
    def get_driver_path(self):
        """Find the correct Edge driver in the specified folder."""
        for root, dirs, files in os.walk(self.driver_location):
            for file in files:
                if file.startswith("msedgedriver") and file.endswith(".exe"):
                    return os.path.join(root, file)
        raise FileNotFoundError("EdgeDriver not found in specified location.")

    def create_driver(self):
        """Create the browser driver."""
        options = webdriver.EdgeOptions()
        options.add_argument("--window-size=1200,900")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        if self.headless:
            options.add_argument("--headless=new")

        service = Service(self.get_driver_path())
        driver = webdriver.Edge(service=service, options=options)
        driver.set_page_load_timeout(60)
        return driver

    # -------------------------------------------------------------------------
    # Utility functions
    # -------------------------------------------------------------------------
    def safe_click(self, browser, locator_type, locator):
        """Click element safely with retries."""
        for attempt in range(self.retry_attempts):
            try:
                element = WebDriverWait(browser, 15).until(
                    EC.element_to_be_clickable((locator_type, locator))
                )
                browser.execute_script("arguments[0].scrollIntoView(true);", element)
                try:
                    element.click()
                except ElementClickInterceptedException:
                    browser.execute_script("arguments[0].click();", element)
                return True
            except TimeoutException:
                print(f"Attempt {attempt+1}: Element {locator} not clickable yet. Retrying...")
                time.sleep(self.retry_delay)
        print(f"Failed to click element after {self.retry_attempts} attempts: {locator}")
        return False

    def wait_for_element(self, browser, locator_type, locator, timeout=30):
        """Wait until element is present."""
        try:
            return WebDriverWait(browser, timeout).until(
                EC.presence_of_element_located((locator_type, locator))
            )
        except TimeoutException:
            print(f"Timeout waiting for element: {locator}")
            return None

    def wait_for_loading_to_disappear(self, browser, timeout=30):
        """Wait for loading overlay/spinner to disappear."""
        try:
            WebDriverWait(browser, timeout).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "dwf-loading-icon"))
            )
            print("✅ Loading overlay disappeared.")
        except TimeoutException:
            print("⚠️ Loading overlay still visible after timeout.")

    def wait_for_page_ready(self, browser, timeout=30):
        """Wait until the document.readyState == 'complete'."""
        WebDriverWait(browser, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("✅ Page fully loaded and ready.")

    def close_dropdown_if_open(self, browser):
        """Close any open dropdowns by clicking outside."""
        try:
            dropdowns = browser.find_elements(By.CLASS_NAME, "open")
            if dropdowns:
                print("Dropdown detected — closing it.")
                browser.execute_script("document.body.click();")
                time.sleep(0.5)
        except Exception:
            pass

    def switch_to_new_tab(self, browser):
        """Switch to newly opened tab and wait for it to load."""
        WebDriverWait(browser, 20).until(lambda d: len(d.window_handles) > 1)
        tabs = browser.window_handles
        browser.switch_to.window(tabs[-1])
        print("🧭 Switched to new tab.")
        self.wait_for_page_ready(browser)
        self.wait_for_loading_to_disappear(browser)

    # -------------------------------------------------------------------------
    # Main workflow
    # -------------------------------------------------------------------------
    def run(self):
        """Main automation logic."""
        try:
            browser = self.create_driver()
            browser.get(self.map_url)
            print("Opened main URL.")

            # Accept cookie popup
            self.safe_click(browser, By.XPATH, "//button[contains(text(), 'Accept') or contains(text(), 'OK')]")

            # Login
            username_field = self.wait_for_element(browser, By.NAME, "username", 10)
            password_field = self.wait_for_element(browser, By.NAME, "password", 10)
            if username_field and password_field:
                username_field.send_keys(self.user_id)
                password_field.send_keys(self.passcode)
                print("Entered credentials.")
                self.safe_click(browser, By.XPATH, "//input[@value='Log in']")
            else:
                print("⚠️ Username or password field not found.")

            # Navigation sequence
            self.safe_click(browser, By.CLASS_NAME, "fis-home")
            self.close_dropdown_if_open(browser)

            # First Accounting click
            self.safe_click(browser, By.XPATH, "//div[text()='Accounting' and contains(@class, 'ng-binding')]")
            self.close_dropdown_if_open(browser)

            # Second Accounting click → opens new tab
            self.safe_click(browser, By.XPATH, "//a[text()='Accounting' and contains(@class, 'ng-binding')]")
            self.switch_to_new_tab(browser)

            # Now wait dynamically until the new page is completely loaded
            self.wait_for_page_ready(browser)
            self.wait_for_loading_to_disappear(browser)

            # Once loaded, find the box and insert data
            input_box = self.wait_for_element(browser, By.XPATH, "//input[@placeholder='Search...']", 25)
            if input_box:
                input_box.send_keys(self.search_value)
                print("✅ Data inserted successfully.")
            else:
                print("❌ Input box not found.")

        except WebDriverException as e:
            print("WebDriver error:", e)
        except Exception as e:
            print("Unexpected error:", e)
        finally:
            try:
                browser.quit()
                print("Browser closed.")
            except Exception:
                pass


# -----------------------------------------------------------------------------
# Run the automation
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    automation = CdrAutomation(browser_name="Edge", headless=False)
    automation.run()
