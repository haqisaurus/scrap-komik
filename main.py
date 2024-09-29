import requests
import uuid
import json
from bs4 import BeautifulSoup
import logging
import re
import openpyxl
import os
from flask import Flask
url_upload = 'https://kisahstory.my.id/api/upload-json'
excel_file = 'database.xlsx'

def download_json_comic(url):
    data = {'uuid': str(uuid.uuid4())}
    chapters = []
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        content = response.content

        # Parse HTML content with BeautifulSoup
        soup = BeautifulSoup(content, 'lxml')

        # Extracting the required information
        title = soup.select_one('#Judul > h1').text
        title = title.replace('Komik', '')
        genre = soup.select_one('#Informasi > table tr:nth-child(3) > td:nth-child(2)').text
        author = soup.select_one("#Informasi > table tr:nth-child(5) > td:nth-child(2)").text
        status = soup.select_one("#Informasi > table tr:nth-child(6) > td:nth-child(2)").text
        how_to_read = soup.select_one("#Informasi > table tr:nth-child(8) > td:nth-child(2)").text
        description = soup.select_one("#Judul > p.desc").text
        thumbnail = soup.select_one("#Informasi > div > img")['src']
        bg = soup.select_one("#Informasi > div > img")['src']

        logging.info(title)

        tags = [tag.text.strip() for tag in soup.select("#Informasi > ul > li a")]

        chapter_rows = soup.select("#Daftar_Chapter tr")[1:]  # Skipping the header row

        for i, chapter_row in enumerate(chapter_rows):
            title_chapter = chapter_row.select_one(".judulseries").text.strip()
            tanggal = chapter_row.select_one(".tanggalseries").text.strip()
            link = "https://komiku.id" + chapter_row.select_one(".judulseries a")['href']
            title_chapter = extract_float(title_chapter)

            images = []
            try:
                chapter_response = requests.get(link)
                chapter_response.raise_for_status()

                chapter_content = chapter_response.content
                chapter_soup = BeautifulSoup(chapter_content, 'lxml')

                images = [{
                    'src': img['src'].strip(),
                    'order': j
                } for j, img in enumerate(chapter_soup.select("#Baca_Komik > img"))]

            except Exception as e:
                logging.error(f"Error fetching chapter: {str(e)}")

            chapters.append({
                'order': title_chapter,
                'tanggal': tanggal,
                'link': link,
                'images': images
            })

        # Reverse the order of chapters
        chapters.reverse()

        data.update({
            'title': title,
            'genre': genre,
            'author': author,
            'status': status,
            'howToRead': how_to_read,
            'description': description,
            'thumbnail': thumbnail,
            'bg': bg,
            'tags': tags,
            'lastChapter': chapters[-1]['order'] if chapters else None,
            'chapters': chapters,
            'url': url
        })

    except Exception as e:
        print(f"Error scraping: {str(e)}")
        raise RuntimeError(f"HTTP Error: {str(e)}") from e

    return data

def extract_float(text):
    # Use regex to extract the numerical float from the string
    result = re.findall(r'\d+\.\d+|\d+', text)
    # Join and convert the list to float, assuming there's a valid number
    return float(result[0]) if result else None

def upload_data(data): 
    # Send the data to the API
    headers = {'Content-Type': 'application/json'}

    try:
        upload_response = requests.post(url_upload, headers=headers, data=json.dumps(data))
        upload_response.raise_for_status()
    except Exception as e:
        print(f"HTTP Error: {str(e)}")
        raise RuntimeError(f"HTTP Error: {str(e)}") from e

def read_excel():
    # Set execution time limits (optional in Python, handled by system)
    input_file_name = os.path.dirname(os.path.abspath(__file__)) +"/"+ excel_file
    result = []
    json_data = None

    try:
        # Load the spreadsheet
        workbook = openpyxl.load_workbook(input_file_name)
        worksheet = workbook['Sheet1']  # Load the specific sheet by name
        
        # Convert the sheet to a list of rows
        data = list(worksheet.iter_rows(values_only=True))

        # Loop through the rows starting from row 10 (index 9)
        for key, value in enumerate(data):
            if key >= 9 and len(result) < 1 and value[1] != 'uploaded':
                logging.info(f"row {key} url:=> {value[0]} ")
                result.append(value)

                # Call the `download_json_comic` function with the first cell (A column value)
                json_data = download_json_comic(value[0])
                upload_data(json_data)

                # Update the 'B' column of the current row to 'uploaded'
                worksheet[f'B{key + 1}'] = 'uploaded'

                # Stop after finding the first row that matches the criteria
                break
        
        # Save the modified spreadsheet
        workbook.save(input_file_name)
    except Exception as e:
        logging.error(f"Error {e}")
        return None
    
    return json_data


app = Flask(__name__)
@app.route('/')
def index():
    result = read_excel()
    if result is None:
        return "Failed"
    result.update({
        "chapters": []
    })
    return result

if __name__ == '__main__':
    app.run(debug=True)