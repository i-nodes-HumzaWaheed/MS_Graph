import time
import msal
import requests
from seleniumbase import Driver, BaseCase
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Microsoft OAuth2 Configuration
CLIENT_ID = "e1c71122-accb-49fe-b425-5d3928ee02dc"
TENANT_ID = "common"  # Use 'common' for personal accounts or specify your Azure tenant
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# Initialize MSAL PublicClientApplication
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Request Device Code
device_flow = app.initiate_device_flow(scopes=SCOPES)
if "message" not in device_flow:
    raise Exception("Failed to get device code. Check your client ID.")

print("Go to:", device_flow["verification_uri"])
print("Enter Code:", device_flow["user_code"])
class automatic(BaseCase):
    def auto(self):
        with Driver(uc=True, headed=True, ad_block=True) as driver:
            # Start Selenium to automate the login process
            driver.get(device_flow["verification_uri"])  # Open the Microsoft login page

            time.sleep(3)  # Wait for page load

            # Find the input field and enter the device code
            code_input = driver.find_element(By.TAG_NAME, "input")
            code_input.send_keys(device_flow["user_code"])
            code_input.send_keys(Keys.RETURN)

            time.sleep(30)

            # Enter Email
            driver.type("//*/input[@type='email' and @name='loginfmt']", "admin@M365x60209898.onmicrosoft.com")
            time.sleep(1)
            try:
                driver.click("//*/button[@type='submit' and text()='Next']")
            except Exception as e:
                driver.click("//*/input[@type='submit' and @value='Next']")
            # email_input = driver.find_element(By.NAME, "loginfmt")
            # email_input.send_keys("dddvdumas@outlook.com")
            # email_input.send_keys(Keys.RETURN)
            try:
                driver.click("//*/button[@type='submit' and text()='Next']")
            except Exception as e:
                try:
                    driver.click("//*/input[@type='submit' and @value='Next']")
                except Exception as e:
                    pass

            time.sleep(30)

            # Enter Password
            driver.type("//*/input[@type='password' and @name='passwd']", "bmU=}q]7(GCP6E+ImqB1$43b1t07%^Wm")
            time.sleep(1)
            try:
                driver.click("//*/button[@type='submit' and text()='Sign in']")
            except Exception as e:
                driver.click("//*/input[@type='submit' and @value='Sign in']")
            # password_input = driver.find_element(By.NAME, "passwd")
            # password_input.send_keys("qecu627oQ")
            # password_input.send_keys(Keys.RETURN)

            time.sleep(3)

            # Handle "Stay Signed In?" prompt if it appears
            try:
                stay_signed_in = driver.find_element(By.ID, "idBtn_Back")
                stay_signed_in.click()  # Click "Yes"
            except:
                print("No 'Stay signed in?' prompt found.")

            print("Login successful. Waiting for token...")
            return
A = automatic()
A.auto()
# Fetch the Access Token
result = app.acquire_token_by_device_flow(device_flow)
if "access_token" in result:
    print("Access Token Acquired:", result["access_token"])

    # Use the token to fetch emails
    headers = {"Authorization": f"Bearer {result['access_token']}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers)
    
    if response.status_code == 200:
        emails = response.json()
        for email in emails["value"]:
            print(email["subject"])
    else:
        print("Error fetching emails:", response.text)
else:
    print("Failed to acquire token:", result.get("error_description", "Unknown error"))
