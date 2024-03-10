from robocorp.tasks import task
from robocorp import browser, log
import json
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF."""
    browser.configure(
        slowmo=1000,
    )
    open_the_intranet_website()
    with log.suppress_variables():
        log_in()
    download_excel_file()
    fill_form_with_excel_data()
    
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

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form."""
    
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def download_excel_file():
    """Downloads the Excel file from the intranet."""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/intranet/SalesData.xlsx", overwrite=True)
