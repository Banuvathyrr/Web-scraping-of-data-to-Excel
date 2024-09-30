# Web-scraping-of-data-to-Excel
**Project: Scrape Financial Data from Moneycontrol to Excel**  
**Overview**  
This project involves scraping financial data (e.g., balance sheets, ratios) from the Moneycontrol website into Excel using the "Get Data" feature. We'll manually extract the appropriate data URL using browser Developer Tools and then feed this URL into Excel to automate data retrieval.

**Steps to Scrape Financial Data from Moneycontrol**  
**Step 1: Identify the URL of the Financial Data**  

**Navigate to Moneycontrol Website:**
- Go to Moneycontrol.
- Search for the financial data page of the company you’re interested in (e.g., Reliance Industries).

**Open Developer Tools (F12):**
Once on the financial page (e.g., ratios, balance sheet), right-click anywhere on the page and select Inspect or press F12 to open Developer Tools.
Go to the Network tab in Developer Tools.

**Capture the Data Request:**
Refresh the webpage with Developer Tools open to capture all network requests.
Look through the network requests for a URL that fetches the table data in JSON or direct HTML format.
Look for URLs that contain information like "financials", "ratios", or "table" (these are likely the ones containing the required data).

**Verify the URL:**
Click on each network request, and examine the Response tab to see if it contains the required data (like financial ratios or balance sheet data).
Once you've identified the correct request that contains the financial data, copy its URL.


**Step 2: Use Excel to Import Data**
Open Excel:

Open a new Excel workbook.
Go to Get Data:

Navigate to the Data tab at the top of Excel.
Click on Get Data > From Other Sources > From Web.
Input the Copied URL:

Paste the URL you copied from Developer Tools into the URL field.
Click OK.
Select the Table to Import:

Excel will attempt to connect to the webpage and may show multiple tables.
Browse through the available tables or data sets to find the one containing the financial data you need.
Select the table and click Load or Transform Data if you need to make any adjustments.
Step 3: Transform Data (Optional)
If you clicked Transform Data, Power Query Editor will open.
Here you can:
Remove unnecessary columns.
Change column headers.
Format data types (e.g., converting text to numbers).
Click Close & Load when finished.
Step 4: Automate Data Refresh
Once the data is imported into Excel, you can set up a schedule to refresh the data:
Go to the Data tab.
Click on Queries & Connections.
Right-click on your query and select Properties.
Set up the refresh frequency to automate the process.
Input the Copied URL:

Paste the URL you copied from Developer Tools into the URL field.
Click OK.
Select the Table to Import:

Excel will attempt to connect to the webpage and may show multiple tables.
Browse through the available tables or data sets to find the one containing the financial data you need.
Select the table and click Load or Transform Data if you need to make any adjustments.
Step 3: Transform Data (Optional)
If you clicked Transform Data, Power Query Editor will open.
Here you can:
Remove unnecessary columns.
Change column headers.
Format data types (e.g., converting text to numbers).
Click Close & Load when finished.
Step 4: Automate Data Refresh
Once the data is imported into Excel, you can set up a schedule to refresh the data:
Go to the Data tab.
Click on Queries & Connections.
Right-click on your query and select Properties.
Set up the refresh frequency to automate the process.
Tips for a Successful Project
Access Restrictions: Moneycontrol may have restrictions in place for preventing scraping, which might return errors like "Access Forbidden". Make sure the URL is accessible directly.
Website Authentication: If the financial data is behind a login, you might need to input credentials to access it. This can be handled by selecting the correct authentication method in the Get Data wizard.
Avoiding Blocks: Some websites may block repeated requests from non-browser clients. To avoid being blocked, do not refresh the data too frequently.
Important Considerations
Website Policy: Scraping financial data may violate the website’s terms of service. Always check the terms and conditions of Moneycontrol before proceeding.
Data Accuracy: Scraped data may change if the website updates its layout or data structure. Regularly verify that your data imports are still accurate.
This method is useful for building a report or dashboard in Excel that contains live or periodically updated financial data, which can then be analyzed or visualized further.
