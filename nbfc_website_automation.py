import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

# Constants
EXCEL_INPUT_PATH = 'nbfcData.xlsx'
EXCEL_OUTPUT_PATH = 'output.xlsx'
SEARCH_QUERY_TEMPLATE = '{} official website'

def read_excel(file_path):
    """Read Excel file and return DataFrame."""
    try:
        return pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        raise

def search_google(query):
    """Perform a Google search and return the first few URLs."""
    search_url = f"https://www.google.com/search?q={query}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    start_time = time.time()
    response = requests.get(search_url, headers=headers)
    search_duration = time.time() - start_time

    soup = BeautifulSoup(response.text, 'html.parser')
    results = []

    for a_tag in soup.find_all('a', href=True):
        href = a_tag['href']
        if 'http' in href and not href.startswith('/url?q='):
            results.append(href.split('?')[0])
        if len(results) >= 5:
            break

    print(f"Google search took {search_duration:.2f} seconds")
    return results

def get_official_website(urls):
    """Check if any of the URLs contain 'official' or 'company'."""
    start_time = time.time()
    for url in urls:
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                title = soup.title.string.lower() if soup.title else ''
                if 'official' in title or 'company' in title:
                    fetch_duration = time.time() - start_time
                    print(f"Fetching {url} took {fetch_duration:.2f} seconds")
                    return url
        except requests.RequestException:
            continue
    return None

def save_progress(df, file_path):
    """Save the current state of the DataFrame to an Excel file."""
    try:
        df.to_excel(file_path, index=False)
        print(f"Progress saved to {file_path}")
    except Exception as e:
        print(f"Error saving progress: {e}")

def process_data(df, save_interval=10):
    """Process the data to find official websites and add to DataFrame."""
    df['Official Website'] = None
    for index, row in df.iterrows():
        print(f"Processing {row['NBFC Name']}...")
        query = SEARCH_QUERY_TEMPLATE.format(row['NBFC Name'])
        urls = search_google(query)
        website = get_official_website(urls)
        df.at[index, 'Official Website'] = website

        # Save progress periodically
        if (index + 1) % save_interval == 0:
            save_progress(df, EXCEL_OUTPUT_PATH)

        time.sleep(1)  # Be polite with requests

    # Final save
    save_progress(df, EXCEL_OUTPUT_PATH)
    return df

def save_to_excel(df, file_path):
    """Save the DataFrame to an Excel file."""
    df.to_excel(file_path, index=False)

def main():
    try:
        # Read the input data
        df = read_excel(EXCEL_INPUT_PATH)
        
        # Check if necessary columns exist
        required_columns = ['Regional Office', 'NBFC Name', 'Address', 'Email ID']
        if not all(col in df.columns for col in required_columns):
            raise ValueError(f"Input Excel file must contain the following columns: {required_columns}")
        
        # Process the data to find official websites
        df_with_websites = process_data(df)
        
        # Save the final result to a new Excel file
        save_to_excel(df_with_websites, EXCEL_OUTPUT_PATH)
        print(f"Results have been saved to {EXCEL_OUTPUT_PATH}")

    except KeyboardInterrupt:
        print("\nProcess interrupted by user.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
