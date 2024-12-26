from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
import pandas as pd
from RPA.PDF import PDF
import time
from RPA.Archive import Archive
import shutil


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
    archive_receipts()
    clean_up()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")
    page = browser.page()
    page.click("text=Order your robot!")
    page.click("text=OK")
    get_orders()


def get_orders():
    """Downloads orders from CSV and processes each order"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    df = pd.read_csv("orders.csv", delimiter=",")
    page = browser.page()

    for i in range(len(df)):
        # Fill form fields
        page.select_option("#head", str(df['Head'][i]))
        page.click(f"input[name='body'][value='{str(df['Body'][i])}']")
        page.fill("input[type='number']", str(df['Legs'][i]))
        page.fill("#address", str(df['Address'][i]))
        page.click("text=Preview")
        page.click("button#order")

        # Handle order button click and check for alerts
        while page.is_visible("div.alert.alert-danger"):
            page.click("button#order")
        
        # Get order number and process receipt
        order_number = df["Order number"][i]
        pdf_path = store_receipt_as_pdf(order_number)
        screenshot_path = screenshot_robot(order_number)
        embed_screenshot_to_receipt(screenshot_path, pdf_path)

        # Prepare for next order
        page.click("button#order-another")
        time.sleep(1)
        page.click("text=OK")


def store_receipt_as_pdf(order_number):
    """Saves the receipt HTML as a PDF"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path


def screenshot_robot(order_number):
    """Takes a screenshot of the ordered robot"""
    page = browser.page()
    screenshot_path = f"output/screenshots/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path


def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embeds the screenshot into the PDF receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot_path, source_path=pdf_path, output_path=pdf_path
    )


def archive_receipts():
    """Archives all receipts into a single ZIP file"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")


def clean_up():
    """Cleans up folders where receipts and screenshots are saved"""
    shutil.rmtree("./output/receipts", ignore_errors=True)
    shutil.rmtree("./output/screenshots", ignore_errors=True)
