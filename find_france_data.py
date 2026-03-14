import json
import os
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


BASE_URL = "https://www.example.com"  # change to the real domain
HREF_JSON_FILE = "href_list.json"
EXCEL_FILE = "francewines.xlsx"


def clean_text(text):
    if not text:
        return ""
    return " ".join(text.split()).strip()


def split_name(full_name):
    full_name = clean_text(full_name)
    if not full_name:
        return "", ""
    parts = full_name.split()
    first_name = parts[0]
    last_name = " ".join(parts[1:]) if len(parts) > 1 else ""
    return first_name, last_name


def get_li_value_from_icon(info_box, icon_name):
    for li in info_box.find_all("li"):
        img = li.find("img")
        if img and icon_name in img.get("src", ""):
            return clean_text(li.get_text(" ", strip=True))
    return ""


def get_first_phone(info_box):
    for li in info_box.find_all("li"):
        img = li.find("img")
        if img and "icon-phone.svg" in img.get("src", ""):
            return clean_text(li.get_text(" ", strip=True))
    return ""


def get_website(info_box):
    for li in info_box.find_all("li"):
        img = li.find("img")
        if img and "icon-link.svg" in img.get("src", ""):
            a_tag = li.find("a", href=True)
            if a_tag:
                return clean_text(a_tag["href"])
            return clean_text(li.get_text(" ", strip=True))
    return ""


def scrape_page(url):
    print(f"Scraping: {url}")

    empty_row = {
        "Company name": "",
        "Email": "",
        "First name": "",
        "Last name": "",
        "Job title": "",
        "country/address": "",
        "Phone": "",
        "Website": "",
    }

    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Failed to fetch {url}: {e}")
        return empty_row

    soup = BeautifulSoup(response.text, "html.parser")
    info_box = soup.find("div", class_="domain-infos")
    if not info_box:
        print(f"No domain-infos section found: {url}")
        return empty_row

    company_tag = info_box.find("strong", class_="domain-name")
    company_name = clean_text(company_tag.get_text()) if company_tag else ""

    address = get_li_value_from_icon(info_box, "icon-address.svg")
    person_name = get_li_value_from_icon(info_box, "icon-name.svg")
    phone = get_first_phone(info_box)
    website = get_website(info_box)

    first_name, last_name = split_name(person_name)

    return {
        "Company name": company_name,
        "Email": "",
        "First name": first_name,
        "Last name": last_name,
        "Job title": "",
        "country/address": address,
        "Phone": phone,
        "Website": website,
    }


def load_links():
    with open(HREF_JSON_FILE, "r", encoding="utf-8") as f:
        hrefs = json.load(f)
    return [urljoin(BASE_URL, href) for href in hrefs]


def find_header_indexes(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if value:
            headers[str(value).strip()] = col
    return headers


def append_row_to_excel(row_data):
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"{EXCEL_FILE} was not found.")

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    headers = find_header_indexes(ws)

    required_headers = [
        "Company name",
        "Email",
        "First name",
        "Last name",
        "Job title",
        "country/address",
        "Phone",
        "Website",
    ]

    for header in required_headers:
        if header not in headers:
            raise ValueError(f'Missing required column header: "{header}"')

    next_row = ws.max_row + 1

    ws.cell(next_row, headers["Company name"], row_data["Company name"])
    ws.cell(next_row, headers["Email"], row_data["Email"])
    ws.cell(next_row, headers["First name"], row_data["First name"])
    ws.cell(next_row, headers["Last name"], row_data["Last name"])
    ws.cell(next_row, headers["Job title"], row_data["Job title"])
    ws.cell(next_row, headers["country/address"], row_data["country/address"])
    ws.cell(next_row, headers["Phone"], row_data["Phone"])
    ws.cell(next_row, headers["Website"], row_data["Website"])

    wb.save(EXCEL_FILE)
    wb.close()


def main():
    links = load_links()

    for link in links:
        row_data = scrape_page(link)
        append_row_to_excel(row_data)
        print(f"Appended row for: {link}")

    print("Done.")


if __name__ == "__main__":
    main()