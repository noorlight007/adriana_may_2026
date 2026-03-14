# https://www.bourgogne-wines.com/

from bs4 import BeautifulSoup
import json

# Read the HTML file
with open("Growers and merchants - Bourgogne wines.html", "r", encoding="utf-8") as f:
    html_content = f.read()

# Parse HTML
soup = BeautifulSoup(html_content, "html.parser")

# Find the container div
container = soup.find("div", id="resultatListeAppellation")

href_list = []

# Extract href from each li > a
if container:
    for li in container.find_all("li"):
        a_tag = li.find("a")
        if a_tag and a_tag.get("href"):
            href_list.append(a_tag["href"])

# Save to JSON file with 4-space indentation
with open("href_list.json", "w", encoding="utf-8") as f:
    json.dump(href_list, f, indent=4, ensure_ascii=False)

print("Saved", len(href_list), "links to href_list.json")