from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
import shutil
from pathlib import Path
import zipfile
import os
import time

@task
def order_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images."""
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    orders = get_orders()
    close_annoying_modal()
    fill_the_form(orders)
    zip_folder("output/results","output/compressed_results.zip")


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
   

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    orders = read_csv_to_table(path= "orders.csv")
    return orders
    

def read_csv_to_table(path, header=True):
    tables = Tables()
    table = tables.read_table_from_csv(
        path=path,
        header=header
    )
    return table

def close_annoying_modal():
    page = browser.page()
    page.click("//*[@class='btn btn-dark']")

def fill_the_form(orders):
    page = browser.page()
    
    for row in orders:
        page.select_option("//*[@class='custom-select']",  str(row["Head"]))
        body = "id-body-"+str(row["Body"])
        page.click("//*[@id='" + body + "']")
        page.fill("(//*[@class='form-control'])[1]", str(row["Legs"]))
        page.fill("(//*[@class='form-control'])[2]", str(row["Address"]))
        page.click("//*[@class='btn btn-primary']")
        

        while page.is_visible("//*[@class='alert alert-danger']"):
                page.click("//*[@class='btn btn-primary']")

        
        pdf = pdf_order_robot(str(row["Order number"]))
        scr = screenshot_robot(str(row["Order number"]))
        embed_screenshot_to_receipt(scr, pdf)

        page.click("//*[@class='btn btn-primary']")
        time.sleep(2)
        close_annoying_modal()


def pdf_order_robot(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    order_results_html = page.locator("//*[@class='col-sm-7']").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(order_results_html, "output/results/order_number_"+order_number+".pdf")
    return  "output/results/order_number_"+order_number+".pdf"
   

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    robot_img = page.query_selector("//*[@id='robot-preview-image']")
    robot_img.screenshot(path="output/results/screenshot_order_number"+order_number+".png")
    return "output/results/screenshot_order_number"+order_number+".png"


def embed_screenshot_to_receipt(screenshot, pdf_file):
    # Adds a watermark image (in this case, a screenshot) to a PDF
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,  # Path to the screenshot image
        source_path=pdf_file,  # Path to the source PDF
        output_path=pdf_file,  # Path where the output PDF will be saved
        coverage=0.2  # Adjusts the scale of the image on the page; adjust as needed
    )    


def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname=arcname)