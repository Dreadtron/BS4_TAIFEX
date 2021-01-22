import os
from datetime import datetime, timedelta

import requests
import openpyxl
import pandas as pd
from bs4 import BeautifulSoup


def generate_url(search_date):
    year = search_date.year
    month = search_date.month
    day = search_date.day
    return f"https://www.taifex.com.tw/cht/3/futContractsDate?queryDate={year}%2F{month}%2F{day}"


def ping_url(url):
    # use .get to find post request(check element ==> network ==> request header) by using "? + key", see <generate_url>
    r = requests.get(url)
    if r.status_code == requests.codes.ok:
        return r.text
    else:
        return None


def parsing(site_text, today):
    soup = BeautifulSoup(site_text, "lxml")  # lxml is another good parser, but need to pip install first
    try:
        table = soup.find("table", class_="table_f")
        table_rows = table.find_all("tr")[3:]
    except AttributeError:
        print(f"No data for {today}")
        table_rows = None

    return table_rows


def row_process(row_data):
    # convert "1,623,302" to int form "1623302"
    converted = [d.replace(",", "") for d in row_data[2:]]
    return row_data[:2] + converted


def table_process(table_rows, today):
    row_product = None  # set up row_product
    row_data_list = []

    if table_rows is None:
        row_data_list = [f"No data for {today}"]                   # row data: weekends
    else:
        for row in table_rows:
            table_cells = row.find_all("td")
            cell_text = [table_cell.text.strip() for table_cell in table_cells]

            # break after finding "期貨小計" because we don't need data from those rows and below
            if cell_text[0] == "期貨小計":
                break

            # rows containing product names have 15 rows instead of 13.
            if len(cell_text) == 15:
                row_product = cell_text[1]
                row_data = row_process(cell_text[1:])              # row data: rows with headers
            else:
                row_data = row_process([row_product] + cell_text)  # row data: rows without headers
            row_data_list.append(row_data)

    return row_data_list


def export_to_excel(export_data, search_date, master_file_name):
    os.makedirs("output", exist_ok=True)          # save to a excel master file
    os.makedirs("output/by_date", exist_ok=True)  # save to individual html file

    # generate or load master file
    writer = pd.ExcelWriter(f"output/{master_file_name}.xlsx", engine='openpyxl')
    if os.path.exists(f"output/{master_file_name}.xlsx"):
        book = openpyxl.load_workbook(f"output/{master_file_name}.xlsx")
        writer.book = book

    # write file
    if len(export_data) != 1:
        filename = search_date.strftime("%Y-%m-%d")
        headers = ['商品名稱', '身份別', '交易多方口數', '交易多方金額', '交易空方口數', '交易空方金額', '交易多空淨口數',
                   '交易多空淨額', '未平倉多方口數', '未平倉多方金額', '未平倉空方口數', '未平倉空方金額', '未平倉淨口數',
                   '未平倉多空淨額']

        df = pd.DataFrame(export_data, columns=headers)  # this kind of Dataframe usually named "df"
        df.to_excel(writer, sheet_name=filename)         # excel file
        writer.save()
        df.to_html(f"output/by_date/{filename}.html")    # html file
        print("工作完成")
    else:
        print("不輸出假日檔案")


def get_data():
    today = datetime.today()
    date_limit = 7  # max. 2 years
    date_count = 0

    # master_file filename
    end_date = today - timedelta(days=(date_limit - 1))
    first_day = today.strftime("%Y-%m-%d")
    last_day = end_date.strftime("%Y-%m-%d")
    master_file_name = f" {first_day} to {last_day}"

    while True:
        search_date = today - timedelta(days=date_count)
        date_count += 1

        if search_date == today - timedelta(days=date_limit):
            print(f"Reach date limit: {date_limit}")
            break

        site_text = ping_url(generate_url(search_date))
        if site_text is not None:
            print(f"Retrieved data from {search_date}")
            table_rows = parsing(site_text, today)
            row_data = table_process(table_rows, today)
            export_to_excel(row_data, search_date, master_file_name)


if __name__ == "__main__":
    get_data()








