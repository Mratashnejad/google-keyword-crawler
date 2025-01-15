import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import urllib.parse

def fetch_google_results(query, start_index=0):
    search_url = "https://www.google.com/search?q=" + urllib.parse.quote(query) + "&start=" + str(start_index)
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    response = requests.get(search_url, headers=headers)
    if response.status_code != 200:
        print(f"Failed to fetch search results for query '{query}'")
        return []
    
    # Debug: Print the raw HTML response to inspect it
    print(f"Raw HTML for query '{query}':\n{response.text[:500]}...")  # Only print the first 500 characters

    # Parse the response content with BeautifulSoup
    soup = BeautifulSoup(response.text, "html.parser")
    links = []
    
    # Find all the search result divs and extract URLs
    for a_tag in soup.find_all('a', href=True):
        href = a_tag['href']
        # Check if it's a valid URL
        if "url?q=" in href:
            link = href.split("url?q=")[1].split("&")[0]
            links.append(link)
    
    # Debug: Print the extracted links
    print(f"Found {len(links)} URLs for query '{query}':")
    for url in links:
        print(f"  {url}")
    
    return links


def search_and_update_excel(file_name):
    try:
        # Load the existing Excel file
        workbook = openpyxl.load_workbook(file_name)
        
        # Select the "Google Search Results" sheet
        sheet = workbook["Google Search Results"]

        # Iterate over rows starting from row 2 (skip header in row 1)
        for row_index in range(2, sheet.max_row + 1):  # Start from row 2
            keyword_cell = sheet.cell(row=row_index, column=1)
            keyword = keyword_cell.value

            # Ensure no hidden spaces or non-visible characters
            if keyword:
                keyword = keyword.strip()

            # Skip empty rows
            if not keyword:
                print(f"Row {row_index} is empty. Skipping...")
                continue
            
            print(f"Searching for: {keyword}")
            
            try:
                urls = []
                # Fetching 40 URLs (4 pages of 10 results each)
                for start_index in range(0, 40, 10):  # Start at 0, then 10, 20, 30
                    result_urls = fetch_google_results(keyword, start_index)
                    urls.extend(result_urls)
                    
                    # Wait to prevent hitting Google too frequently
                    print(f"Waiting for a few seconds before fetching the next page of results...")
                    time.sleep(5)  # Adjust the delay as needed (e.g., 5 seconds)
                
                # Print each found URL in the console
                for url in urls:
                    print(f"Found URL: {url}")
                
                # Write URLs into the sheet starting from column B
                for col_index, url in enumerate(urls, start=2):
                    sheet.cell(row=row_index, column=col_index).value = url
                
                print(f"Results written for keyword: {keyword}")
            except Exception as e:
                print(f"Error fetching results for '{keyword}': {e}")
        
        # Save changes to the same file
        workbook.save(file_name)
        print(f"All updates saved to {file_name}")
    except FileNotFoundError:
        print(f"File {file_name} not found. Please check the file path.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Specify the Excel file name
    excel_file = "GoogleSearchResults.xlsx"
    search_and_update_excel(excel_file)
