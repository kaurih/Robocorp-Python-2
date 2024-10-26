from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time


@task
def order_robots_from_RobotSpareBin():
    browser.configure(slowmo=500)

    #Orders robots from RobotSpareBin Industries Inc.
    #Saves the order HTML receipt as a PDF.
    #Saves the screenshot of the ordered robot.
    #Embeds the screenshot of the robot to the PDF receipt.
    #Creates ZIP archive of the receipts and the images.

    open_robot_order_website()

    orders = get_orders()
    for index, row in enumerate(orders):
        try:
            fill_the_form(row)
            store_receipt_as_pdf(str(index))
        except Exception as e:
            print(f"Error processing order {index + 1}: {e}")
            break  # Exit on failure
        finally:
            # click order-another
            if index < len(orders) - 1:
                prepare_next_order()

    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()
    fetch_csv()


def close_annoying_modal():
    page = browser.page()
    try:
        if page.locator("button:text('OK')").is_visible():
            page.click("button:text('OK')")
    except Exception:
        pass  # continue if no modal


def fetch_csv():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def get_orders():
    orders_file = Tables()
    return orders_file.read_table_from_csv("orders.csv")


def fill_the_form(row):
    page = browser.page()
    
    # head
    page.select_option("select#head", row["Head"])
    
    # body
    page.click(f'//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{row["Body"]}]/label')

    # legs
    page.fill("input[placeholder='Enter the part number for the legs']", row["Legs"])
    page.fill("input#address", row["Address"])

    page.click("button#preview")

    # click order (and again on error)
    while True:
        page.click("button#order")
        
        if page.locator("div.alert.alert-danger").is_visible():
            print("retrying submission")
            time.sleep(0.5)
        else:
            break
    
    store_receipt_as_pdf(row["Order number"])


def prepare_next_order():
    page = browser.page()
    page.click("#order-another")
    close_annoying_modal()


def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_path = f"output/receipts/receipt_{order_number}.pdf"
    pdf.html_to_pdf(receipt, pdf_path)

    #screenshot and embed
    screenshot_robot(order_number)
    embed_screenshot(order_number)
    return pdf_path


def embed_screenshot(order_number):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=f"output/receipts/order_{order_number}.png",
        source_path=f"output/receipts/receipt_{order_number}.pdf",
        output_path=f"output/receipts/receipt_{order_number}.pdf",
        coverage=0.15 #adjust pic size
    )


def screenshot_robot(order_number):
    page = browser.page()
    page.locator("#robot-preview-image").screenshot(path=f"output/receipts/order_{order_number}.png")


def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")
