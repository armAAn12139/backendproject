import pandas as pd
from googleapiclient.discovery import build
from tqdm import tqdm
import logging

logging.basicConfig(filename='nbfc_website_scraper.log', level=logging.INFO)


def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return None


def find_official_website(nbfc_name):
    api_key = AIzaSyA_MKz3eFx0HDV_Ie6eUjVBf8hphdnaWXc
    #cse_id = add here
    search_service = build('customsearch', 'v1', developerKey=api_key)
    search_query = f'{nbfc_name} official website'
    results = search_service.cse().list(q=search_query, cx=cse_id).execute()
    website_url = None
    for result in results['items']:
        if result['link'].startswith('http'):
            website_url = result['link']
            break
    return website_url


def validate_and_store_website(df, nbfc_name, website_url):
    if website_url:
        df.loc[df['NBFC Name'] == nbfc_name, 'Official Website'] = website_url
    return df


def main():
    file_path = 'NBFCsandARCs10012023 (5).XLSX'
    df = read_excel_file("C:\\Users\\Dell\\IdeaProjects\\python.py\\project.py")
    if df is None:
        logging.error("Error reading Excel file")
        return

    pbar = tqdm(total=len(df), desc='Processing NBFCs')

    for index, row in df.iterrows():
        nbfc_name = row['NBFC Name']
        website_url = find_official_website(nbfc_name)
        df = validate_and_store_website(df, nbfc_name, website_url)
        pbar.update(1)

    pbar.close()

    output_file_path = 'NBFCs_with_official_websites.xlsx'
    df.to_excel(output_file_path, index=False)

    logging.info("Program completed successfully")


if __name__ == '__main__':
    main()
