from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
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
    browser.configure(
        slowmo=2000,
    )
    download_the_orders_file()
    open_the_intranet_robot_order_website()
    log_in()
    click_order_robot()
    close_annoying_modal_popup()
    read_csv_file_as_table()
    archive_receipts_to_zip()


def archive_receipts_to_zip():
    """Archives all receipt PDFs into a single ZIP file."""
    archiver = Archive()
    archiver.archive_folder_with_zip(
        folder="output/Receipts",
        archive_name="output/robot_order_receipts.zip"
    )


def download_the_orders_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    print("Orders file downloaded.")
   

def read_csv_file_as_table():
    """Read data from csv as table"""

    tables = Tables()
    table = tables.read_table_from_csv("orders.csv")
    for row in table:
        fill_form_from_row(row)
   

def open_the_intranet_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")
    

def click_order_robot():
    page = browser.page()
    page.click("text=Order your robot!")

def close_annoying_modal_popup():
    page = browser.page()
    page.click("button:has-text('OK')")

def fill_form_from_row(row):

    page = browser.page()

    print(page.url)
    # Head dropdown
    page.select_option("#head", str(row["Head"]))

    # Body radio button
    page.click(f"input[name='body'][value='{row['Body']}']")

    # Legs input
    page.fill("input[placeholder='Enter the part number for the legs']", str(row["Legs"]))

    # Address textarea
    page.fill("#address", row["Address"])

    # Retry logic for order submission
    max_retries = 3
    for attempt in range(max_retries):
        page.click("button:has-text('Order')")
        # Wait for either error message or order completion
        page.wait_for_timeout(1000)  # Wait 1 second for UI update
        if not page.is_visible(".alert-danger"):
            break
        if attempt == max_retries - 1:
            print(f"Order failed after {max_retries} attempts for order {row['Order number']}")
            return
    store_receipt_as_pdf(row["Order number"])

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#order-completion").inner_html()

    pdf = PDF()
    pdf_file = f"output/Receipts/order_{order_number}.pdf"
    pdf.html_to_pdf(sales_results_html, pdf_file)
    screenshot_file = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot_file, pdf_file)

    page.click("button:has-text('Order another robot')")
    close_annoying_modal_popup()

def screenshot_robot(order_number):
    """Takes a screenshot of the robot image"""
    page = browser.page()
    robot_image = page.locator("#robot-preview-image")
    page.wait_for_timeout(2500)  # Wait 2.5 second for UI update
    screenshot_path = f"output/Robots_image/robot_{order_number}.png"
    robot_image.screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot of the robot to the PDF receipt"""
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )