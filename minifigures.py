from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from collections import namedtuple
import pandas
import requests
from fpdf import FPDF
from PIL import Image
from typing import Tuple
from tqdm import tqdm
import math
import re
import os


# local globals
CHROME_DRIVER_PATH = "C:/Users/Alex/Desktop/chromedriver/chromedriver-win64/chromedriver.exe"
REQUESTS_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}
IMG_PATH = "img"
PDF_NAME = "my_minifigures.pdf"
EXCEL_PATH = "my_minifigures.xlsx"
EXCEL_SHEET_NAME = "Minifigures"
EXCEL_LINK_COLUMN_NAME = "Link"
EXCEL_QUANTITY_COLUMN_NAME = "Quantity"

# page-related globals
PAGE_WAIT_DELAY = 0.3
USED_FIGURE_TABLE_ID = 1
AVG_PRICE_PATTERN = r"Avg Price: UAH ([\d,]+\.\d+)"
MINIFIGURE_NAME_ELEMENT = "item-name-title"
MINIFIGURE_PRICE_ELEMENT = "pcipgSummaryTable"
MINIFIGURE_IMG_ELEMENT = "_idImageMain"
MINIFIGURE_NAME_MAX_LENGTH = 46

# data types
MinifigureInputData = namedtuple("MinifigureInputData", ["link", "quantity"])
MinifigureWebData = namedtuple("MinifigureWebData", ["name", "price", "quantity", "image"])


def read_excel(file_path: str) -> list[MinifigureInputData]:
    print("Reading excel data...")
    minifigure_list = []
    excel_data = pandas.read_excel(file_path, sheet_name=EXCEL_SHEET_NAME)
    for _, row in excel_data.iterrows():
        link = row[EXCEL_LINK_COLUMN_NAME]
        quantity = row.get(EXCEL_QUANTITY_COLUMN_NAME, None)
        minifigure_list.append(MinifigureInputData(link=link, quantity=quantity))

    return minifigure_list


def normalize_minifigure_name(name: str) -> str:
    return re.sub(r"[\'/]", "", name)


def fetch_minifigures_data(input_list: list[MinifigureInputData]) -> list[MinifigureWebData]:
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    if not os.path.exists(IMG_PATH):
        os.makedirs(IMG_PATH)

    result_list = []

    for input_element in tqdm(input_list, desc="Fetching web data"):
        driver.get(input_element.link)
        driver.implicitly_wait(PAGE_WAIT_DELAY)

        fetched_name = driver.find_element(By.ID, MINIFIGURE_NAME_ELEMENT).text
        minifigure_name = normalize_minifigure_name(fetched_name)
        if len(minifigure_name) > MINIFIGURE_NAME_MAX_LENGTH:
            minifigure_name = minifigure_name[:MINIFIGURE_NAME_MAX_LENGTH] + "..."

        price_tables = driver.find_elements(By.CLASS_NAME, MINIFIGURE_PRICE_ELEMENT)
        table_text = price_tables[USED_FIGURE_TABLE_ID].text
        minifigure_price = float(re.search(AVG_PRICE_PATTERN, table_text).group(1).replace(",", ""))

        image_element = driver.find_element(By.ID, MINIFIGURE_IMG_ELEMENT)
        image_url = image_element.get_attribute("src")
        response = requests.get(image_url, headers=REQUESTS_HEADERS)

        if response.status_code == 200:
            image_bytes = BytesIO(response.content)
            image = Image.open(image_bytes)
            rgb_im = image.convert("RGB")
            minifigure_img = f"{IMG_PATH}/{minifigure_name}.jpg"
            rgb_im.save(minifigure_img)
        else:
            raise ValueError(f"Failed to get image of {input_element.link}")

        result_list.append(MinifigureWebData(name=minifigure_name, price=minifigure_price, quantity=input_element.quantity, image=minifigure_img))

    driver.quit()

    return result_list


def sort_minifigure_list(input_list: list[MinifigureWebData]) -> Tuple[float, list[MinifigureWebData]]:
    print("Sorting minifigure list by price...")
    minifigures_sorted = sorted(input_list, key=lambda x: x.price, reverse=True)
    total_value = sum(figure.price * (1 if figure.quantity is None or math.isnan(figure.quantity) else int(figure.quantity)) for figure in minifigures_sorted)

    return total_value, minifigures_sorted


def create_pdf_document(total_value: float, input_list: list[MinifigureWebData]):
    print("Creating pdf document...")
    pdf = FPDF()
    pdf.add_page()

    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt=f"My LEGO StarWars minifigures total value: UAH {total_value}", ln=True, align="C")

    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)

    element_index = 1

    for minifigure_data in input_list:
        if pdf.get_y() > 260:
            pdf.add_page()

        current_x = pdf.get_x()
        current_y = pdf.get_y()

        pdf.cell(120, 10, f"{element_index}. {minifigure_data.name}", ln=0)

        price_value = f"UAH {minifigure_data.price}" if not minifigure_data.quantity or math.isnan(minifigure_data.quantity) else f"UAH {minifigure_data.price} x{int(minifigure_data.quantity)}"
        pdf.cell(20, 10, price_value, ln=0)

        pdf.image(minifigure_data.image, x=current_x + 166, y=current_y, w=20, h=20)

        pdf.ln(22)
        element_index += 1

    pdf.output(PDF_NAME)
    print("PDF generated successfully!")


def main():
    input_list = read_excel(EXCEL_PATH)
    web_data_list = fetch_minifigures_data(input_list)
    total_value, sorted_list = sort_minifigure_list(web_data_list)
    create_pdf_document(total_value, sorted_list)


if __name__ == "__main__":
    main()
