from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
from dotenv import load_dotenv
import logging


# Load credentials from .env file
load_dotenv('credentials.env')
USERNAME = os.getenv("BOT_USERNAME")
PASSWORD = os.getenv("BOT_PASSWORD")

# Logging setup
logging.basicConfig(
    filename="robot_spare_bin.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

class Bot:

    def process_sales_data(self):
        try:
            open_the_intranet_website()
            log_in()
            download_excel_file()
            fill_form_with_excel_data()
            collect_results()
            export_as_pdf()
        except Exception as e:
            logging.error(f"Error in processing: {e}")
            self.errors.append(str(e))
            self.handle_error()
        else:
            logging.info("Data processed successfully")
        finally:
            log_out()

@task
def robot_spare_bin_python():
    """Main task to process sales data"""
    bot = Bot()
    bot.process_sales_data()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.configure(slowmo=100)
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Logs in using secured credentials"""
    page = browser.page()
    page.fill("#username", USERNAME)
    page.fill("#password", PASSWORD)
    page.click("button:text('Log in')")

def fill_and_submit_sales_form(sales_rep):
    """Fills in sales data and clicks the Submit button"""
    try:
        page = browser.page()
        page.fill("#firstname", sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget", str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")
        logging.info(f"Successfully submitted data for {sales_rep['First Name']} {sales_rep['Last Name']}")
    except Exception as e:
        logging.error(f"Error submitting data for {sales_rep['First Name']} {sales_rep['Last Name']}: {e}")
        raise

def download_excel_file():
    """Downloads Excel file from the given URL"""
    http = HTTP()
    try:
        http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
        logging.info("Excel file downloaded successfully")
    except Exception as e:
        logging.error(f"Failed to download Excel file: {e}")
        raise

def fill_form_with_excel_data():
    """Reads data from Excel and fills in the sales form"""
    excel = Files()
    try:
        excel.open_workbook("SalesData.xlsx")
        worksheet = excel.read_worksheet_as_table(header=True)
        for row in worksheet:
            try:
                fill_and_submit_sales_form(row)
            except Exception:
                logging.warning(f"Skipping row: {row}")
        excel.close_workbook()
    except Exception as e:
        logging.error(f"Error processing Excel file: {e}")
        raise

def collect_results():
    """Takes a screenshot of the page"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")
    logging.info("Screenshot saved successfully")

def export_as_pdf():
    """Exports the sales results to a PDF file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
    logging.info("PDF export completed")

def log_out():
    """Logs out of the application"""
    page = browser.page()
    page.click("text=Log out")
    logging.info("Logged out successfully")
