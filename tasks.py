from robocorp.tasks import task
from robocorp.workitems import WorkItems
from robocorp.browser import Browser
from RPA.Excel.Files import Files as Excel
import re

@task
def main():
    # Initialize work items and browser
    work_items = WorkItems()
    browser = Browser()
    excel = Excel()
    
    # Open the Excel file to store data
    excel.create_workbook("news_data.xlsx")
    excel.rename_worksheet("Sheet", "Data")
    excel.append_row(["Title", "Date", "Description", "Picture Filename", "Count", "Contains Money"])
    
    # Process each input work item
    for item in work_items.inputs:
        search_phrase = item.payload["search_phrase"]
        news_category = item.payload["news_category"]
        number_of_months = item.payload["number_of_months"]
        
        # Implement web scraping logic here
        # Example: browser.goto("https://example-news-site.com")
        # Perform search, navigate, and extract data
        
        # For each news item extracted, process and store data in Excel
        # Example row to append:
        # excel.append_row([title, date, description, picture_filename, count, contains_money])
        
    # Save and close the Excel workbook
    excel.save("output/news_data.xlsx")
    excel.close()
    
    # Close the browser
    browser.close()

if __name__ == "__main__":
    main()