# ──────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Credentials and
# company-specific references have been removed.
# In production this loads from environment
# variables set via Streamlit secrets or .env.
# See README.md for context.
# ──────────────────────────────────────────────

import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyotp

# Load credentials from environment variables
# Required keys: LONG_EMAIL, NETSUITE_PASSWORD, NETSUITE_KEY (TOTP secret)
LONG_EMAIL          = os.getenv("LONG_EMAIL")
NETSUITE_PASSWORD   = os.getenv("NETSUITE_PASSWORD")
NETSUITE_KEY        = os.getenv("NETSUITE_KEY")

# NetSuite 2FA field IDs can vary by account configuration.
# Update these if your instance uses different element IDs.
NETSUITE_2FA_INPUT_ID  = os.getenv("NETSUITE_2FA_INPUT_ID",  "uif60_input")
NETSUITE_2FA_SUBMIT_ID = os.getenv("NETSUITE_2FA_SUBMIT_ID", "uif76")


def netsuite_login(driver, log_message, url=None):
    """
    Performs login to NetSuite with email, password, and TOTP 2FA.

    Args:
        driver:      Selenium WebDriver instance
        log_message: Callable that accepts a string — used for timestamped logging
        url:         Optional URL to navigate to before login.
                     If None, assumes driver is already on the login page.

    Raises:
        ValueError:  If required environment variables are missing
        Exception:   If login fails at any step; saves a screenshot on failure
    """
    if not all([LONG_EMAIL, NETSUITE_PASSWORD, NETSUITE_KEY]):
        log_message("Missing required environment variables (LONG_EMAIL, NETSUITE_PASSWORD, NETSUITE_KEY)")
        raise ValueError("Required credentials not set in environment.")

    log_message("Starting NetSuite login")

    if url:
        driver.get(url)
        log_message(f"Navigated to: {url}")

    try:
        # ── Email ──────────────────────────────────────────────────────────
        log_message("Waiting for email field...")
        email_field = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "email"))
        )
        email_field.clear()
        email_field.send_keys(LONG_EMAIL)
        log_message("Entered email")

        # ── Password ───────────────────────────────────────────────────────
        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(NETSUITE_PASSWORD)
        log_message("Entered password")

        # ── Submit ─────────────────────────────────────────────────────────
        driver.find_element(By.ID, "login-submit").click()
        log_message("Submitted credentials")

        # ── TOTP 2FA ───────────────────────────────────────────────────────
        log_message("Waiting for 2FA field...")
        twofa_field = WebDriverWait(driver, 25).until(
            EC.visibility_of_element_located((By.ID, NETSUITE_2FA_INPUT_ID))
        )
        totp = pyotp.TOTP(NETSUITE_KEY)
        code = totp.now()
        twofa_field.clear()
        twofa_field.send_keys(code)
        log_message("Entered 2FA code")

        driver.find_element(By.ID, NETSUITE_2FA_SUBMIT_ID).click()
        log_message("Submitted 2FA")

        # ── Confirm redirect away from login page ──────────────────────────
        time.sleep(4)
        WebDriverWait(driver, 20).until(
            lambda d: "app/login" not in d.current_url
        )
        log_message("Login successful — redirected to dashboard")

    except Exception as e:
        log_message(f"NetSuite login error: {e}")
        driver.save_screenshot("netsuite_login_error.png")
        raise
