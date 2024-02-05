from csv import DictReader

from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
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
    browser.configure(slowmo=100)
    open_robot_order_website()
    get_orders()
    handle_orders()
    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def handle_orders():
    with open("orders.csv", encoding="UTF-8") as file:
        csv_reader = DictReader(file)

        for row in csv_reader:
            fill_the_form(row=row)


def close_annoying_modal():
    page = browser.page()
    page.click(selector="text=OK")


def fill_the_form(row: dict):
    close_annoying_modal()
    page = browser.page()
    page.select_option(selector="#head", value=row["Head"])
    page.click(selector=f"#id-body-{row['Body']}")
    page.fill(selector="//input[@type=\"number\"]", value=row["Legs"])
    page.fill(selector="#address", value=row["Address"])
    page.click(selector="text=Preview")
    page.click(selector="#order")
    page.wait_for_load_state(state="networkidle")

    while page.is_visible(selector="#order"):
        try:
            page.mouse.wheel(delta_y=100, delta_x=0)
            page.click(selector="#order", timeout=20, force=True)
            page.wait_for_load_state(state="networkidle")
        except Exception:
            break

    pdf_path = receipt_to_pdf(order_number=row["Order number"])
    jpeg_path = screenshot_robot(order_number=row["Order number"])
    embed_screenshot_to_receipt(screenshot=jpeg_path, pdf_file=pdf_path)
    page.click(selector="#order-another")


def receipt_to_pdf(order_number: int):
    pdf_path = f"output/receipts/order_{order_number}.pdf"
    page = browser.page()
    receipt_html = page.locator(selector="#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(content=receipt_html, output_path=pdf_path)
    return pdf_path


def screenshot_robot(order_number: int):
    jpeg_path = f"output/screenshots/order_{order_number}.jpeg"
    page = browser.page()
    page.locator(selector="#robot-preview-image").screenshot(type="jpeg", path=jpeg_path)
    return jpeg_path


def embed_screenshot_to_receipt(screenshot: str, pdf_file: str):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip(folder="output/receipts/", archive_name="receipts.zip")
