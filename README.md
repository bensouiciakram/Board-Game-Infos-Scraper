# 🎲 Board Game Infos Scraper

This Python script extracts detailed information about board games listed on [BoardGameGeek](https://boardgamegeek.com) using game links from an Excel file. It collects data such as player counts, designers, categories, and images, then exports the results to a new Excel workbook.

---

## 📌 Features

- ✅ Scrapes data from **BoardGameGeek** using XPath and JSON parsing.
- 🖼️ Fetches up to 3 popular images per game via the GeekDo API.
- 📄 Processes two Excel sheets simultaneously and appends extracted data.
- 📦 Outputs a clean, structured Excel file with the enriched data.
- 🛡️ Includes basic error handling and logging for traceability.

---

## 📂 Input File

The script expects an Excel file named: `With-Without Numbers Game Links.xlsx`
It should contain two sheets:
- **Table 1**
- **Table 2**

Each sheet must include a column titled `BGG Link` containing URLs of BoardGameGeek game pages.

---

## 📤 Output

After execution, the script generates:`final_output.xlsx`
Containing enriched versions of both tables, including new fields like:
- Game Name
- Game Box (cover image URL)
- Minimum & Maximum Players
- Playing Time
- Game Weight (complexity)
- Designer, Artist, Publisher
- Game Type, Category, Mechanisms
- Description
- Images (top 3 hot ones)

---

## 🔧 Installation

Install the required dependencies:

```bash
pip install -r requirements.txt
```
## 🚀 Usage
Just run the script:

```bash
python games_infos_scraper.py
```
You'll see progress and logs in your console as each game is processed.

## 🧠 How It Works 
The script:
1. Loads links from the Excel file.
2. For each game link:
   * Downloads the HTML.
   * Extracts JSON-embedded metadata.
   * Parses game attributes via `nested_lookup`.
   * Pulls image URLs via GeekDo's API.
3. Writes results to `final_output.xlsx`.

