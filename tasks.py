from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem


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
    open_browser()
    download_file_and_fill_form()
    archive_pdf_files()
    cleanup_folder()


def open_browser():
    """Open the browser and log in"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def download_file_and_fill_form():
    """Download the orders file"""
    http = HTTP()
    http.download(
        "https://robotsparebinindustries.com/orders.csv", overwrite=True)
    csv = Tables()
    orders = csv.read_table_from_csv("orders.csv")
    for row in orders:
        fill_one_form(row)


def fill_one_form(order):
    """Fill one form"""
    page = browser.page()
    page.click("button:text('OK')")
    page.select_option("#head", order['Head'])
    page.click(f"#id-body-{order['Body']}")
    page.get_by_placeholder(
        "Enter the part number for the legs").fill(order['Legs'])
    page.fill("#address", order['Address'])
    page.click("#preview")
    page.locator(
        "#robot-preview-image").screenshot(path=f"output/{order['Order number']}.png")
    receipt = page.locator("#receipt")
    while receipt.count() <= 0:
        page.click("#order")
    receipt.screenshot(path=f"output/receipt_{order['Order number']}.png")
    files_to_add = [
        f"output/receipt_{order['Order number']}.png", f"output/{order['Order number']}.png"]
    add_files_to_pdf(files_to_add, order)
    page.click("#order-another")


def add_files_to_pdf(files, order):
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=files, target_document=f"output/order_{order['Order number']}.pdf")


def archive_pdf_files():
    """Archive all pdf files inside the output folder"""
    lib = Archive()
    lib.archive_folder_with_zip(
        './output', './output/orders.zip', include='*.pdf')


def cleanup_folder():
    """Remove png and pdf files from the output folder."""
    fs = FileSystem()
    files = fs.list_files_in_directory("./output")
    for file in files:
        if file.name.endswith(".pdf") or file.name.endswith(".png"):
            fs.remove_file(file)
            print(file, "deleted.")
