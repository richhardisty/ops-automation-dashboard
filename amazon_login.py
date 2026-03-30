# ──────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Credentials and
# company-specific references have been removed.
# In production this loads from a .env file via
# python-dotenv. See README.md for context.
# ──────────────────────────────────────────────

import os
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyotp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# Load credentials from environment variables
# In production: loaded from .env via load_dotenv()
# Required keys: EMAIL, AMAZON_PASSWORD, AMAZON_KEY (TOTP secret)
EMAIL            = os.getenv("EMAIL")
AMAZON_PASSWORD  = os.getenv("AMAZON_PASSWORD")
AMAZON_KEY       = os.getenv("AMAZON_KEY")

# Vendor Central account name shown in the account selector
# Replace with your account's display name
VENDOR_ACCOUNT_NAME = os.getenv("VENDOR_ACCOUNT_NAME", "Your Vendor Account Name")


def amazon_login(driver, log_message):
    """
    Performs login to Amazon Vendor Central with email, password, and TOTP 2FA.

    Args:
        driver:      Selenium WebDriver instance
        log_message: Callable that accepts a string — used for timestamped logging

    Raises:
        ValueError:  If required environment variables are missing
        Exception:   If login fails at any step
    """
    if not all([EMAIL, AMAZON_PASSWORD, AMAZON_KEY]):
        log_message("Missing required environment variables (EMAIL, AMAZON_PASSWORD, AMAZON_KEY)")
        raise ValueError("Required credentials not set in environment.")

    log_message("Starting Amazon Vendor Central login")

    while True:
        try:
            # ── Email ──────────────────────────────────────────────────────
            driver.find_element(By.ID, "ap_email").send_keys(EMAIL)
            log_message("Entered email")

            try:
                password_field = driver.find_element(By.ID, "ap_password")
            except Exception:
                password_field = None

            if not password_field:
                driver.find_element(By.ID, "continue").click()
                log_message("Clicked Continue")
                time.sleep(1)

            # ── Password ───────────────────────────────────────────────────
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ap_password"))
            ).send_keys(AMAZON_PASSWORD)
            log_message("Entered password")

            driver.find_element(By.ID, "signInSubmit").click()
            time.sleep(3)
            log_message("Submitted credentials")

            # ── TOTP 2FA ───────────────────────────────────────────────────
            totp = pyotp.TOTP(AMAZON_KEY)
            driver.find_element(By.ID, "auth-mfa-otpcode").send_keys(totp.now())
            log_message("Entered 2FA code")
            driver.find_element(By.ID, "auth-signin-button").click()
            time.sleep(3)
            log_message("Submitted 2FA")

            # ── Account selector ───────────────────────────────────────────
            # Vendor Central shows an account picker after login if multiple
            # accounts are associated with the credential.
            driver.find_element(
                By.XPATH,
                f'//button/span[text()="{VENDOR_ACCOUNT_NAME}"]'
            ).click()

            # The "Select account" button uses a Shadow DOM web component
            driver.execute_script("""
                const btn = document.querySelector('kat-button[label="Select account"]');
                if (btn) {
                    const inner = btn.shadowRoot.querySelector('button');
                    if (inner) inner.click();
                }
            """)
            log_message("Selected vendor account")
            time.sleep(3)

            # ── Dismiss tour dialog if present ─────────────────────────────
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        '//button[contains(@class, "take-tour-dialog-content-ctas-tertiary")'
                        ' and contains(text(), "Maybe later")]'
                    ))
                ).click()
                log_message("Dismissed tour dialog")
            except Exception:
                pass  # Dialog not present — continue

            # ── Check for error page ───────────────────────────────────────
            try:
                err = driver.find_element(
                    By.XPATH, '//li[contains(text(), "Error, please try again.")]'
                )
                if err.is_displayed():
                    log_message("Error page detected — retrying")
                    driver.refresh()
                    continue
            except Exception:
                pass

            log_message("Login successful")
            break

        except Exception as e:
            log_message(f"Login error: {e} — retrying")
            driver.refresh()
            continue


if __name__ == "__main__":
    """
    Standalone test mode — initializes a Chrome WebDriver, logs in,
    and pauses for inspection. Useful for verifying credentials independently.
    """

    def _log(message):
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}")

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=chrome_options
    )

    try:
        driver.get("https://vendorcentral.amazon.com")
        _log("Navigated to Amazon Vendor Central")
        amazon_login(driver, _log)
        _log("Login completed successfully")
        input("Press Enter to close the browser...")
    except Exception as e:
        _log(f"Login failed: {e}")
        raise
    finally:
        driver.quit()
        _log("WebDriver closed")
