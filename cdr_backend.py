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
        self.map_url = "https://example.com/login"
        self.headless = headless
        self.browser_name = browser_name
        self.retry_attempts = 3
        self.retry_delay = 2

    def get_driver_path(self):
        """Locate the browser driver executable."""
        for root, dirs, files in os.walk(self.driver_location):
            for file in files:
                if file.startswith("msedgedriver") and file.endswith(".exe"):
                    return os.path.join(root, file)
        raise FileNotFoundError("EdgeDriver not found in specified location.")

    def create_driver(self):
        """Create and configure the WebDriver instance."""
        if self.browser_name.lower() == "edge":
            options = webdriver.EdgeOptions()
        else:
            options = webdriver.ChromeOptions()

        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--window-size=1200,900")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        driver_path = self.get_driver_path()
        service = Service(driver_path)

        if self.browser_name.lower() == "edge":
            return webdriver.Edge(service=service, options=options)
        else:
            return webdriver.Chrome(service=service, options=options)

    def safe_click(self, browser, locator_type, locator):
        """Safely click an element with retry logic."""
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
                print(f"Attempt {attempt + 1}: Element {locator} not clickable yet. Retrying...")
                time.sleep(self.retry_delay)
        print(f"Failed to click element after {self.retry_attempts} attempts: {locator}")
        return False

    def wait_for_element(self, browser, locator_type, locator, timeout=20):
        """Wait for an element to appear on the page."""
        try:
            return WebDriverWait(browser, timeout).until(
                EC.presence_of_element_located((locator_type, locator))
            )
        except TimeoutException:
            print(f"Element not found within timeout: {locator}")
            return None

    def wait_for_loading_to_disappear(self, browser, timeout=30):
        """Wait until loading spinner disappears."""
        try:
            WebDriverWait(browser, timeout).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "dwf-loading-icon"))
            )
            print("Loading overlay disappeared.")
        except TimeoutException:
            print("Loading overlay still present after timeout.")

    def close_dropdown_if_open(self, browser):
        """Close dropdown if any open element detected."""
        try:
            dropdown = browser.find_elements(By.CLASS_NAME, "open")
            if dropdown:
                print("Dropdown detected. Closing...")
                browser.execute_script("document.body.click();")
                time.sleep(1)
        except Exception:
            pass

    def switch_to_new_tab(self, browser):
        """Switch to newly opened tab."""
        WebDriverWait(browser, 10).until(lambda d: len(d.window_handles) > 1)
        tabs = browser.window_handles
        browser.switch_to.window(tabs[-1])
        print("Switched to new tab.")
        WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        self.wait_for_loading_to_disappear(browser)

    def wait_until_page_ready(self, browser, timeout=30):
        """Wait until document.readyState == 'complete'."""
        WebDriverWait(browser, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("Page fully loaded and ready.")

    def run(self):
        """Main automation workflow."""
        try:
            browser = self.create_driver()
            browser.get(self.map_url)

            # Accept cookies if present
            self.safe_click(browser, By.XPATH, "//button[contains(text(), 'Accept') or contains(text(), 'OK')]")

            # Login
            username_field = self.wait_for_element(browser, By.NAME, "username", timeout=10)
            password_field = self.wait_for_element(browser, By.NAME, "password", timeout=10)

            if username_field and password_field:
                username_field.send_keys(self.user_id)
                password_field.send_keys(self.passcode)
                print("Entered username and password.")
            else:
                print("Username/password fields not found. Proceeding to login anyway.")

            self.safe_click(browser, By.XPATH, "//input[@value='Log in']")

            # Navigate to Accounting
            self.safe_click(browser, By.CLASS_NAME, "fis-home")
            self.close_dropdown_if_open(browser)

            self.safe_click(browser, By.XPATH, "//div[text()='Accounting' and contains(@class, 'ng-binding')]")
            self.close_dropdown_if_open(browser)

            # Click second Accounting link and wait for full page load
            self.safe_click(browser, By.XPATH, "//a[text()='Accounting' and contains(@class, 'ng-binding')]")
            self.switch_to_new_tab(browser)
            self.wait_until_page_ready(browser)
            self.wait_for_loading_to_disappear(browser)

            # Now the page is fully loaded → safe to interact with elements
            print("Now ready to insert data into the box.")

            # Example: wait for input box and send value
            input_box = self.wait_for_element(browser, By.XPATH, "//input[@placeholder='Search...']", timeout=20)
            if input_box:
                input_box.send_keys(self.search_value)
                print("Data inserted successfully.")
            else:
                print("Input box not found.")

        except WebDriverException as e:
            print("WebDriver error:", e)
        except Exception as e:
            print("Error during automation:", e)
        finally:
            try:
                browser.quit()
                print("Browser closed.")
            except Exception:
                pass


if __name__ == "__main__":
    automation = CdrAutomation(browser_name="Edge", headless=False)
    automation.run()
