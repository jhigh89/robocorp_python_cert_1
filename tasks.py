from robocorp.tasks import task
from robocorp import browser, log
import json
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF."""
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    with log.suppress_variables():
        log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()
    
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
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def collect_results():
    """Take a screenshot of the page."""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the page as a PDF."""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """Log out from the intranet."""
    page = browser.page()
    page.click("text=Log out")