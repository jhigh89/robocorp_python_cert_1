from robocorp.tasks import task
from robocorp import browser, log
import json

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF."""
    browser.configure(
        slowmo=1000,
    )
    open_the_intranet_website()
    with log.suppress_variables():
        log_in()
    fill_and_submit_sales_form()
    
def load_config(config_path):
    try:
        with open(config_path, 'r') as file:
            config = json.load(file)
            return config
    except FileNotFoundError:
        print("Config file not found.")
    except json.JSONDecodeError:
        print("Error decoding JSON from the config file.")

def open_the_intranet_website():
    """Open the intranet website by navigating to the URL."""
    browser.goto("https://robotsparebinindustries.com/intranet/")

def log_in():
    """Log in to the intranet."""
    page = browser.page()

    config = load_config('config.json')
    login_credentials = config.get('login_credentials', {})
    username = login_credentials.get('username', '')
    password = login_credentials.get('password', '')
    page.fill("#username", username)
    page.fill("#password", password)
    page.click("button:text('Log in')")

def fill_and_submit_sales_form():
    """Fills in the sales data and clicks the 'Submit' button."""
    page = browser.page()

    page.fill("#firstname", "John")
    page.fill("#lastname", "Smith")
    page.fill("#salesresult", "123")
    page.select_option("#salestarget", "10000")
    page.click("text=Submit")
    