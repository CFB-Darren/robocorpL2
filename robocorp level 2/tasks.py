from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    orders = get_orders()
    open_browser()
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
    archive_receipts()
    

def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=['Order number', 'Head', 'Body', 'Legs', 'Address']
    )
    return orders

def open_browser():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def fill_the_form(order):
    page = browser.page()
    head_names = {
        "1" : "Roll-a-thor head",
        "2" : "Peanut crusher head",
        "3" : "D.A.V.E head",
        "4" : "Andy Roid head",
        "5" : "Spanner mate head",
        "6" : "Drillbit 2000 head"
    }
    head_number = order['Head']
    page.select_option("#head", head_names.get(head_number))
    page.click("#id-body-" + str(order['Body']))
    page.fill("input[placeholder='Enter the part number for the legs']", str(order['Legs']))
    page.fill("#address", str(order['Address']))

    preview_robot()
    submit_robot()
    pdf = store_receipt_as_pdf(order['Order number'])
    screenshot = screenshot_robot(order['Order number'])
    embed_screenshot_to_receipt(screenshot, pdf)
    page.click("#order-another")

def preview_robot():
    page = browser.page()
    page.click("#preview")

def submit_robot():
    page = browser.page()
    while True:
        page.click("#order")
        retry_error = page.query_selector("#order-another")
        if retry_error:
            break

def store_receipt_as_pdf(order_number):
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/PDF/" + order_number + ".pdf")
    return "output/PDF/" + order_number + ".pdf"

def screenshot_robot(order_number):
    page = browser.page()
    page.locator("#robot-preview-image").screenshot(path="output/Screenshot/" + order_number + ".png")
    return "output/Screenshot/" + order_number + ".png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file, output_path=pdf_file)

def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip("output/PDF", "output/pdf.zip")