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
    WebDriverException,
)


class CdrAutomation:
    def __init__(self, browser_name: str, headless=False):
        self.user_id = "your_user_id"
        self.passcode = "your_password"
        self.search_value = "Mountain Private Equity Partners"
        self.driver_location = f"C:\\Drivers\\{browser_name}"
        self.map_url = "https://example.com/login"  # <-- replace with real URL
        self.headless = headless
        self.browser_name = browser_name
        self.retry_attempts = 3
        self.retry_delay = 2

    # -------------------------------------------------------------------------
    # Driver setup
    # -------------------------------------------------------------------------
    def get_driver_path(self):
        for root, dirs, files in os.walk(self.driver_location):
            for file in files:
                if file.startswith("msedgedriver") and file.endswith(".exe"):
                    return os.path.join(root, file)
        raise FileNotFoundError("EdgeDriver not found in specified location.")

    def create_driver(self):
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
    # Utilities
    # -------------------------------------------------------------------------
    def safe_click(self, browser, locator_type, locator, description=""):
        """Click an element safely. If it fails, return False so automation can stop."""
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
                print(f"✅ Clicked: {description or locator}")
                return True
            except TimeoutException:
                print(f"⚠️ Attempt {attempt+1}: {description or locator} not clickable. Retrying...")
                time.sleep(self.retry_delay)
        print(f"❌ Failed to click: {description or locator}")
        return False

    def wait_for_element(self, browser, locator_type, locator, timeout=30):
        try:
            return WebDriverWait(browser, timeout).until(
                EC.presence_of_element_located((locator_type, locator))
            )
        except TimeoutException:
            return None

    def wait_for_page_ready(self, browser, timeout=30):
        WebDriverWait(browser, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

    def wait_for_loading_to_disappear(self, browser, timeout=30):
        """Wait for loading spinner (if any) to disappear."""
        try:
            WebDriverWait(browser, timeout).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "dwf-loading-icon"))
            )
        except TimeoutException:
            print("⚠️ Loading spinner did not disappear (may not exist).")

    def switch_to_new_tab(self, browser):
        """Switch to the newest tab."""
        WebDriverWait(browser, 20).until(lambda d: len(d.window_handles) > 1)
        browser.switch_to.window(browser.window_handles[-1])
        self.wait_for_page_ready(browser)
        self.wait_for_loading_to_disappear(browser)
        print("🧭 Switched to new tab and waited for load.")

    # -------------------------------------------------------------------------
    # Main workflow
    # -------------------------------------------------------------------------
    def run(self):
        browser = None
        try:
            browser = self.create_driver()
            browser.get(self.map_url)
            print("🌐 Opened login page.")

            # Accept cookies if popup exists
            self.safe_click(browser, By.XPATH, "//button[contains(text(), 'Accept') or contains(text(), 'OK')]", "Accept Cookies")

            # Try to find username/password fields
            username_field = self.wait_for_element(browser, By.NAME, "username", 5)
            password_field = self.wait_for_element(browser, By.NAME, "password", 5)

            # Logic: if login fields available → enter creds, else just click Login
            if username_field and password_field:
                username_field.send_keys(self.user_id)
                password_field.send_keys(self.passcode)
                print("👤 Credentials entered.")
                clicked = self.safe_click(browser, By.XPATH, "//input[@value='Log in']", "Login Button")
                if not clicked:
                    print("❌ Login button could not be clicked. Stopping automation.")
                    return
            else:
                print("⚠️ Username/password fields not found — trying direct login.")
                clicked = self.safe_click(browser, By.XPATH, "//input[@value='Log in']", "Login Button (Direct)")
                if not clicked:
                    print("❌ Direct login button not clickable. Stopping automation.")
                    return

            # Wait for main page to load
            self.wait_for_page_ready(browser)
            self.wait_for_loading_to_disappear(browser)

            # Go to Accounting
            if not self.safe_click(browser, By.CLASS_NAME, "fis-home", "Home Button"):
                print("❌ Home button click failed. Stopping automation.")
                return

            if not self.safe_click(browser, By.XPATH, "//div[text()='Accounting' and contains(@class, 'ng-binding')]", "Accounting (Menu)"):
                print("❌ Accounting menu click failed. Stopping automation.")
                return

            if not self.safe_click(browser, By.XPATH, "//a[text()='Accounting' and contains(@class, 'ng-binding')]", "Accounting (Link)"):
                print("❌ Accounting link click failed. Stopping automation.")
                return

            # Switch tab and wait for full load
            self.switch_to_new_tab(browser)
            print("📄 Accounting page opened successfully.")

            # Once loaded, wait for input and enter data
            input_box = self.wait_for_element(browser, By.XPATH, "//input[@placeholder='Search...']", 20)
            if input_box:
                input_box.send_keys(self.search_value)
                print("✅ Search value entered successfully.")
            else:
                print("❌ Input box not found after Accounting page loaded.")

        except WebDriverException as e:
            print("⚠️ WebDriver error:", e)
        except Exception as e:
            print("⚠️ Unexpected error:", e)
        finally:
            if browser:
                try:
                    browser.quit()
                    print("🧹 Browser closed.")
                except Exception:
                    pass


# -------------------------------------------------------------------------
# Entry point
# -------------------------------------------------------------------------
if __name__ == "__main__":
    automation = CdrAutomation(browser_name="Edge", headless=False)
    automation.run()
