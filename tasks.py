from robocorp.tasks import task
from RPA.HTTP import HTTP
from robocorp import browser
from RPA import Tables
from RPA.PDF import PDF
from RPA.Browser.Selenium import Selenium
from RPA import FileSystem
from RPA.Archive import Archive
from selenium import webdriver


http = HTTP()
table = Tables.Tables()
pdf = PDF()
rpa_browser = Selenium()
archive = Archive()
file_sys = FileSystem.FileSystem()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    create_directories()
    open_robot_order_website()
    download_order_file()
    get_orders()
    archive_receipts()

def create_directories():
    dir_exist = FileSystem.Path('output').is_dir()
    if dir_exist:
        FileSystem.FileSystem().remove_directory('output', True)
    FileSystem.Path('output/screenshots').mkdir(parents=True, exist_ok=True)
    FileSystem.Path('output/receipts').mkdir(parents=True, exist_ok=True)
    file_sys.create_file("output/output.robolog")
    

def open_robot_order_website():
    rpa_browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")

def download_order_file():
    http.download("https://robotsparebinindustries.com/orders.csv", "output/orders.csv", True, False, True)

def get_orders():
    orders = table.read_table_from_csv("output/orders.csv")
    for order in orders:
        fill_the_form(order)

def fill_the_form(order):
    close_annoying_modal()
    rpa_browser.wait_until_element_is_enabled(f'//select[@id="head"]/option[@value="{order["Head"]}"]', 20)
    rpa_browser.click_element(f'//select[@id="head"]/option[@value="{order["Head"]}"]')
    rpa_browser.click_element(f'//input[@id="id-body-{order["Body"]}"]')
    rpa_browser.input_text("//*[@placeholder='Enter the part number for the legs']", order["Legs"])
    rpa_browser.input_text("//*[@id='address']", order["Address"])
    rpa_browser.click_element("//button[@id='preview']")
    rpa_browser.wait_until_element_is_visible("//button[@id='order']")
    rpa_browser.click_element("//button[@id='order']")
    try:
        rpa_browser.wait_until_element_is_visible("//*[@id='receipt']")
    except Exception as e:
        rpa_browser.click_element("//button[@id='order']")
    store_receipt_as_pdf(order["Order number"])

def close_annoying_modal():
    try:
        rpa_browser.wait_until_element_is_visible('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]', 20)
        rpa_browser.click_element('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
    except Exception as e:
        try:
            rpa_browser.click_element("//button[@id='order-another']")
        except Exception as e:
            pass
        close_annoying_modal()

def store_receipt_as_pdf(order_number):
    try:
        rpa_browser.wait_until_element_is_visible("//*[@id='receipt']")
        html = rpa_browser.find_element("//*[@id='receipt']").get_attribute("innerHTML")
        rpa_browser.screenshot("//div[@id='robot-preview-image']", f"output/screenshots/{order_number}.png")
        pdf.html_to_pdf(html, f"output/receipts/{order_number}.pdf")
        embed_screenshot_to_receipt(f"output/screenshots/{order_number}.png", f"output/receipts/{order_number}.pdf")
        rpa_browser.wait_until_element_is_visible("//button[@id='order-another']", 20)
        rpa_browser.click_element("//button[@id='order-another']")
    except Exception as e:
        rpa_browser.wait_until_element_is_visible("//button[@id='order']", 20)
        rpa_browser.click_element("//button[@id='order']")
        

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )

def archive_receipts():
    archive.archive_folder_with_tar('./output/receipts', 'output/receipts.zip', recursive=True)









