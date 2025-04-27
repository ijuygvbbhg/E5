#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft 365 Admin Center并检查E5订阅有效期。
- 在GitHub Actions上运行时，浏览器和驱动程序由工作流安装。
- 环境变量 `MS_E5_ACCOUNTS` 从 GitHub Secrets 读取: email-password&email2-password2...
- (可选) `sendNotify.py` 用于发送通知，需要配置相应的 Secrets。
"""
import os
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

# --- Optional Notification Setup ---
# Ensure sendNotify.py is in your repository if you use this
try:
    from sendNotify import send
    # Check for necessary notification secrets passed as environment variables
    # Add checks relevant to your sendNotify.py implementation
    # Example: if not os.environ.get('PUSH_PLUS_TOKEN'): print("Warning: PUSH_PLUS_TOKEN secret not set.")
except ImportError:
    print("通知文件 sendNotify.py 未找到，将仅打印到控制台。")
    def send(title, content):
        print(f"--- {title} ---")
        print(content)
        print("--- End Notification ---")
# --- End Notification Setup ---

List = [] # To store output messages

# --- Configuration ---
LOGIN_URL = 'https://admin.microsoft.com/'
SUBSCRIPTIONS_URL = 'https://admin.microsoft.com/Adminportal/Home?source=applauncher#/subscriptions'
TARGET_SUBSCRIPTION_NAME = "Microsoft 365 E5" 
# Adjust if your E5 subscription name is slightly different

# --- Helper Function ---
def get_webdriver():
    options = webdriver.ChromeOptions()
    # Crucial options for GitHub Actions/headless environments
    options.add_argument("--headless") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") 
    options.add_argument("--window-size=1920,1080")
    # Use a common user agent
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36") 
    
    # In GitHub Actions with apt install, chromedriver should be in PATH
    try:
       # Let Selenium find chromedriver in PATH
       driver = webdriver.Chrome(options=options) 
       List.append("  - WebDriver 初始化成功。")
       return driver
    except WebDriverException as e:
       List.append(f"!! 错误：无法初始化WebDriver: {e}")
       List.append("!! 请检查工作流中的 ChromeDriver 安装步骤。")
       return None
    except Exception as e:
       List.append(f"!! 错误：初始化WebDriver时发生意外错误: {e}")
       return None


def check_e5_expiry(username, password):
    """Logs into Microsoft Admin Center and checks E5 subscription expiry."""
    List.append(f"开始检查账号: {username}")
    driver = get_webdriver()
    if not driver:
        List.append(f"!! 检查失败: {username} (WebDriver 初始化失败)")
        return 

    try:
        driver.get(LOGIN_URL)
        # Increased wait time for potentially slow cloud environments
        wait = WebDriverWait(driver, 45) 

        # --- Login Step 1: Enter Email ---
        try:
            email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
            email_field.send_keys(username)
            # Use JavaScript click as a fallback if direct click fails
            next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", next_button)
            # next_button.click() # Direct click sometimes fails
            List.append("  - 输入邮箱并点击下一步")
        except (NoSuchElementException, TimeoutException) as e:
            List.append(f"!! 错误：找不到邮箱输入框或超时。页面可能更改。 {e}")
            driver.save_screenshot(f"error_email_input_{username}.png") 
            return # Stop check for this user

        time.sleep(random.uniform(3, 5)) # Wait for password or other prompts

        # --- Login Step 2: Enter Password ---
        try:
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
            time.sleep(0.5)
            password_field.send_keys(password)
            signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", signin_button)
            # signin_button.click() 
            List.append("  - 输入密码并点击登录")
        except (NoSuchElementException, TimeoutException) as e:
            # Check if it's asking for password again (common if email format was slightly off or domain federated)
            try:
                if driver.find_element(By.ID, "i0118").is_displayed():
                   List.append("!! 警告: 似乎仍在密码页面，密码可能错误或登录流程异常。")
                else: raise NoSuchElementException # Re-raise if not the password field
            except NoSuchElementException:
                List.append(f"!! 错误：找不到密码输入框或登录按钮。密码错误或页面结构更改。 {e}")
            driver.save_screenshot(f"error_password_input_{username}.png")
            return

        # --- Login Step 3: Handle "Stay signed in?" (KMSI) ---
        try:
            kmsi_button_no = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back")) # The "No" button
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
            # kmsi_button_no.click() 
            List.append("  - 处理 '保持登录状态?' -> 否")
        except TimeoutException:
            List.append("  - 未出现 '保持登录状态?' 弹窗 (或已超时)，继续...")
            # It's possible login failed silently before this, or the page flow changed.
            # Check if we are on an expected page (like the admin dashboard)
            if "admin.microsoft.com" not in driver.current_url:
                 List.append("!! 警告: 未出现KMSI弹窗，且当前URL不是Admin Center。登录可能失败。")
                 driver.save_screenshot(f"error_post_login_url_{username}.png")
                 # Consider returning here if strict login check is needed
        except NoSuchElementException as e:
            List.append(f"!! 错误：无法找到 '保持登录状态?' 按钮。 {e}")
            driver.save_screenshot(f"error_kmsi_button_{username}.png")
            # Continue cautiously

        # --- Navigate to Subscriptions Page ---
        List.append("  - 尝试导航到订阅页面...")
        time.sleep(random.uniform(4, 7)) # Give time for potential redirects
        
        try:
            driver.get(SUBSCRIPTIONS_URL)
            # Wait for a reliable element on the subscriptions page.
            # This XPath looks for the main content area of the 'Your products' page
            # Adjust based on UI changes or language differences (e.g., '产品')
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-is-scrollable='true']")))
            List.append("  - 成功导航到订阅页面")
            time.sleep(random.uniform(2, 4)) # Let dynamic content load
        except TimeoutException:
            List.append("!! 错误：导航到订阅页面超时或找不到预期元素。登录失败或页面结构更改。")
            driver.save_screenshot(f"error_nav_subscriptions_{username}.png")
            return
        except
