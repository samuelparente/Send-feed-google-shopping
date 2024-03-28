# Send-feed-google-shopping
 
This app is a C# windows form .NET application designed to automate the process of generating a product feed for Google Shopping. This application fetches product data from an ERP database, calculates profit margins, and creates an XML file conforming to Google's specifications. The generated XML file can then be uploaded to an FTP server for use in Google Shopping campaigns.

### Features

- **Automated Feed Generation**: FeedSend connects to an ERP database to retrieve product information, including prices and tax rates, then calculates profit margins.
- **XML Feed Creation**: Based on the retrieved data, the application creates an XML file with each product's information, including the calculated profit margin.
- **FTP Upload**: The generated XML file can be uploaded to an FTP server for use in Google Shopping campaigns.

### How it Works

1. **Configuration**: Users can configure the application by providing necessary details such as ERP database connection information, FTP server details, and URLs for fetching product data.
   
2. **ERP Database Connection**: FeedSend establishes a connection to the ERP database using the provided credentials.

3. **Data Retrieval**: The application retrieves product information from the ERP database, including prices and tax rates.

4. **Profit Margin Calculation**: FeedSend calculates the profit margin for each product based on cost prices and tax rates.

5. **XML Feed Generation**: Using the calculated profit margins, the application creates an XML file in the required format for Google Shopping.

6. **FTP Upload**: The generated XML file is then uploaded to the specified FTP server location.

### Usage

1. **Setup**: Before running the application, ensure that the necessary configuration files (`assets/configs.config` and `assets/configsDb.config`) are correctly set up with the required details.

2. **Running the Application**:
   - Launch the application (`feedSend.exe`).
   - Click on the "Edit" button to enter configuration details for the ERP database and FTP server.
   - Once configured, click on the "Save" button to save the settings.
   - The application will fetch data from the ERP database, calculate profit margins, generate the XML feed, and upload it to the FTP server.

### Configuration

- **ERP Database Configuration**: Configure the connection details for the ERP database in `assets/configsDb.config`.
- **FTP Server Configuration**: Provide FTP server details in `assets/configs.config`.

### Requirements

- .NET Framework
- Access to an ERP database with product information
- FTP server for uploading the generated XML feed

### Author

Samuel Parente
