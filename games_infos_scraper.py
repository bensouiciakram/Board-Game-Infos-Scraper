# importation part ---------------------------------------------------------------------------------------#
import pandas as pd
from parsel import Selector
import requests
from re import findall
import json
from nested_lookup import nested_lookup
from pprint import pprint
import traceback
import logging
from typing import Any, Dict, List, Optional, Union


# global variable and initialisations -------------------------------------------------------------------#
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s:%(message)s'
)

game_url_template = 'https://boardgamegeek.com/boardgame/{}'
images_url_template = 'https://api.geekdo.com/api/images?ajax=1&galleries%5B%5D=game&nosession=1&objectid={}&objecttype=thing&showcount=17&size=crop100&sort=hot'

workbook = pd.ExcelFile('With-Without Numbers Game Links.xlsx')
df1 = workbook.parse(sheet_name='Table 1')
df2 = workbook.parse(sheet_name='Table 2')

fields = [
    'Game Name', 'Game Box', 'Minimum Players', 'Maximum Players',
    'Recommended Players', 'Playing Time', 'Weight', 'Alternative Name',
    'Designer', 'Artist', 'Short Description', 'Publisher', 'Type',
    'Category', 'Mechanisms', 'Family', 'Images'
]

# Add missing columns to the DataFrames
for field in fields:
    df1[field] = df2[field] = pd.Series(dtype='object')


# helper functions ------------------------------------------------------------------------------------------#
def data_object_extractor(selector: Selector) -> Dict[str, Any]:
    """
    Extract the JSON object embedded in a <script> tag on the BGG game page.

    Args:
        selector (Selector): Parsel Selector for the HTML content.

    Returns:
        Dict[str, Any]: Parsed JSON object containing game data.
    """
    raw_data_location = selector.xpath('//script[contains(text(),"geekitemPreload")]').get()
    regex = r'geekitemPreload = (\{[\s\S]+?\});'
    return json.loads(findall(regex, raw_data_location)[0])


def get_description_raw_text(description_source: str) -> str:
    """
    Extract raw text content from the game description HTML.

    Args:
        description_source (str): HTML content of the description.

    Returns:
        str: Plain text description.
    """
    description_selector = Selector(text=description_source)
    return description_selector.xpath('string(.)').get()


def image_extractor(url: str) -> List[str]:
    """
    Fetch up to 3 hot images of the game from the GeekDo API.

    Args:
        url (str): BGG game URL.

    Returns:
        List[str]: List of image URLs.
    """
    source = requests.get(images_url_template.format(url.split('/')[-2])).json()
    return nested_lookup('imageurl', source)[:3]


def safe_list_get(l: List[Any]) -> Union[Any, str]:
    """
    Safely get the first item from a list, return an empty string if list is empty.

    Args:
        l (List[Any]): List to access.

    Returns:
        Any or str: First item or empty string.
    """
    try:
        return l[0]
    except IndexError:
        return ''


def game_infos_extractor(url: str) -> Dict[str, Any]:
    """
    Scrape and extract detailed information about a board game from its BGG page.

    Args:
        url (str): The BGG game URL.

    Returns:
        Dict[str, Any]: A dictionary of extracted game fields and values.
    """
    response = requests.get(url)
    selector = Selector(text=response.text)
    source = data_object_extractor(selector)
    return {
        'BGG Link': url,
        'Game Name': nested_lookup("name", source)[0],
        'Game Box': selector.xpath('//meta[@property="og:image"]/@content').get(),
        'Minimum Players': int(nested_lookup("minplayers", source)[0]),
        'Maximum Players': int(nested_lookup("maxplayers", source)[0]),
        'Recommended Players': safe_list_get(nested_lookup('max', nested_lookup("best", source))),
        'Playing Time': int(nested_lookup("minplaytime", source)[0]),
        'Weight': round(float(nested_lookup("avgweight", source)[0]), 2),
        'Alternative Name': nested_lookup('alternatename', source)[0],
        'Designer': ','.join(nested_lookup('name', nested_lookup("boardgamedesigner", source))),
        'Artist': ','.join(nested_lookup('name', nested_lookup("boardgameartist", source))),
        'Short Description': get_description_raw_text(nested_lookup('description', source)[0]),
        'Publisher': nested_lookup('name', nested_lookup("boardgamepublisher", source))[0],
        'Type': ','.join([game_type for game_type in nested_lookup("subdomain", source) if game_type]),
        'Category': ','.join(nested_lookup('name', nested_lookup("boardgamecategory", source))),
        'Mechanisms': ','.join(nested_lookup('name', nested_lookup("boardgamemechanic", source))),
        'Family': ','.join(nested_lookup('name', nested_lookup("boardgamefamily", source))),
        'Images': image_extractor(url)
    }


# main functionality --------------------------------------------------------------------------------------------#
if __name__ == '__main__':
    # Process first table
    for index, row in df1.iterrows():
        logging.info(f'{index} - Extracting the information from the URL: {row["BGG Link"]}')
        item = game_infos_extractor(row['BGG Link'])
        df1.loc[index] = pd.Series(item)
        pprint(f'The extracted infos: {item}\n\n')

    # Process second table
    for index, row in df2.iterrows():
        logging.info(f'{index} - Extracting the information from the URL: {row["BGG Link"]}')
        item = game_infos_extractor(row['BGG Link'])
        df2.loc[index] = pd.Series(item)
        pprint(f'The extracted infos: {item}\n\n')

    # Write to Excel
    writer = pd.ExcelWriter('final_output.xlsx')
    df1.to_excel(writer, 'Table 1')
    df2.to_excel(writer, 'Table 2')
    writer.close()
