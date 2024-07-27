# NBFC Website Automation

## Project Overview

This project automates the process of finding official websites for a list of Non-Banking Financial Companies (NBFCs). The input data is sourced from an Excel file containing information about various NBFCs. The script performs a Google search for each NBFC to locate their official website and then saves the results in a new Excel file. The script now includes functionality to periodically save progress and handle interruptions.

## Objective
 
To build a Python script that:
- Reads an Excel file containing information about NBFCs.
- Performs a Google search for each NBFC to find their official website.
- Saves the results (official website URLs) into a new Excel file.
- Handles interruptions by periodically saving progress.

## Setup of Virtual Environment

1. Create a virtual environment:
    ```sh
    python -m venv venv
    ```

2. Activate the virtual environment:
   - On Windows:
     ```sh
     venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```sh
     source venv/bin/activate
     ```

## Libraries Used

- `requests`: For making HTTP requests to Google and fetching web pages.
- `beautifulsoup4`: For parsing HTML content and extracting data from web pages.
- `pandas`: For handling Excel file operations and managing data in DataFrames.
- `time`: For measuring the duration of various operations and adding delays between requests.

## Code Overview

### Constants

- `EXCEL_INPUT_PATH`: Path to the input Excel file containing NBFC details.
- `EXCEL_OUTPUT_PATH`: Path where the output Excel file with official website URLs will be saved.
- `SEARCH_QUERY_TEMPLATE`: Template string for forming Google search queries.

### Functions

1. **`read_excel(file_path)`**
   - Reads the input Excel file into a DataFrame.
   - Parameters: `file_path` (str) - Path to the Excel file.
   - Returns: DataFrame containing the data from the Excel file.

2. **`search_google(query)`**
   - Performs a Google search for the given query and extracts the top URLs.
   - Parameters: `query` (str) - Search query.
   - Returns: List of URLs obtained from the search results.

3. **`get_official_website(urls)`**
   - Checks if any of the URLs contain 'official' or 'company' in their title.
   - Parameters: `urls` (list of str) - List of URLs to check.
   - Returns: The first URL that contains 'official' or 'company' in the title, or `None` if no such URL is found.

4. **`process_data(df)`**
   - Processes the input DataFrame to find official websites for each NBFC and adds the results to the DataFrame.
   - Periodically saves progress to handle interruptions.
   - Parameters: `df` (DataFrame) - DataFrame containing the NBFC data.
   - Returns: DataFrame with an additional column for official websites.

5. **`save_to_excel(df, file_path)`**
   - Saves the DataFrame to an Excel file.
   - Parameters: `df` (DataFrame) - DataFrame to be saved.
   - Parameters: `file_path` (str) - Path to the output Excel file.

6. **`main()`**
   - Main function to execute the entire process: read the data, process it to find official websites, and save the results.
   - Handles potential interruptions and ensures that the script can resume processing without losing progress.

## Execution

1. Ensure you have the required libraries installed:
    ```sh
    pip install requests beautifulsoup4 pandas openpyxl
    ```

2. Place your input Excel file (e.g., `nbfcData.xlsx`) in the same directory as the script or update the `EXCEL_INPUT_PATH` to the correct file path.

3. Execute the script:
    ```sh
    python script_name.py
    ```

4. Check the output file (e.g., `output.xlsx`) for the results.

## Points of Improvement

- **Handling Search Results**: The script retrieves only the first 5 URLs from Google. Consider increasing this limit or handling pagination if necessary.
- **Error Handling**: Improve error handling for cases where no valid URL is found or the web page is inaccessible.
- **Proxies and Rate Limiting**: Implement proxies and rate limiting to manage cases where Google might block requests due to excessive queries in a short period.
- **Dynamic User-Agent**: Rotate user-agent strings to avoid detection and blocking by Google.

## Edge Cases

- **Empty or Malformed Excel Files**: Ensure that the input file is correctly formatted and contains the necessary columns.
- **Unresponsive Websites**: Some websites may be down or slow to respond, which can affect the script's performance.
- **Captcha Challenges**: Google may prompt Captchas if it detects automated scraping, which would require manual intervention.

## Important Notes

- **Google Search**: The script relies on Google's search results, which may vary based on location and other factors. Ensure compliance with Google's Terms of Service when scraping search results.
- **Network Issues**: Ensure stable network connectivity as the script makes multiple HTTP requests.
- **Performance**: Depending on the number of NBFCs and the speed of web pages, the process might take some time.

## Conclusion

This script provides a streamlined method to automate the process of finding official websites for NBFCs using web scraping techniques. By integrating `requests` and `BeautifulSoup` with `pandas` for data management, it achieves its objective efficiently while allowing for improvements and handling edge cases. The addition of periodic progress saving ensures that data is not lost in case of interruptions.

