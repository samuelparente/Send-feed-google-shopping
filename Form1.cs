// Author: Samuel Parente
// Date: March 28, 2024
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Xml.Linq;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Diagnostics;
using System.Globalization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading;

namespace feedSend
{
    public partial class Form1 : Form
    {
        string configFilePath = "assets/configs.config";
        string dbConfigFilePath = "assets/configsDb.config";

        private System.Windows.Forms.Timer hourlyTimer;

        public Form1()
        {
            InitializeComponent();

            // Subscribe to the Resize event
            this.Resize += Form1_Resize;
            // Subscribe to the MouseDoubleClick event
            notifyIcon1.MouseDoubleClick += notifyIcon_MouseDoubleClick;
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            Show();

            // Initial state of components
            initStat();

            // Configure the timer
            hourlyTimer = new System.Windows.Forms.Timer();
            hourlyTimer.Interval = 3600000; // 1 hour in milliseconds
            hourlyTimer.Tick += HourlyTimer_Tick;

            // First call to update feed
            await LoadDataAsync();

            // Start the timer to call the function every hour
            hourlyTimer.Start();
        }

        #region Trayicon

        // Minimize to tray icon
        private void Form1_Resize(object sender, EventArgs e)
        {
            // If the form is minimized
            // hide it from the task bar
            // and show the system tray icon (represented by the NotifyIcon control)
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }

        // Show form again on tray icon click
        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        #endregion

        #region Functions

        private async void HourlyTimer_Tick(object sender, EventArgs e)
        {
            // Call the LoadDataAsync function
            await LoadDataAsync();
        }

        // Background tasks
        private async Task LoadDataAsync()
        {
            button1.Enabled = false;
            editBtn.Enabled = false;

            if (populateInfo())
            {
                SqlConnection connection = await Task.Run(() => connectDb());
                if (connection != null)
                {
                    if (await Task.Run(() => doFeed(connection)))
                    {
                        string ftpServer = textBox2.Text;
                        string ftpUsername = textBox3.Text;
                        string ftpPassword = textBox4.Text;
                        string localFilePath = "profit.xml";
                        string remoteDirectory = "/import/images/SKU/";
                        await Task.Run(() => UploadFileToFtp(ftpServer, ftpUsername, ftpPassword, localFilePath, remoteDirectory));
                        connection.Close();
                    }

                    connection.Close();

                    UpdateLabel("Disconnected from Sage DB.", Color.Black, 3);
                }
            }

            button1.Enabled = true;
            editBtn.Enabled = true;
        }

        // Initial state
        private void initStat()
        {
            toolStripStatusLabel1.Text = " ";
            toolStripStatusLabel2.Text = " ";
            toolStripStatusLabel3.Text = " ";
            toolStripStatusLabel4.Text = " ";

            editBtn.Enabled = true;
            readConfigBtn.Enabled = false;
            clrBtn.Enabled = false;
            saveBtn.Enabled = false;
            cancelBtn.Enabled = false;
            button1.Enabled = true;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
        }

        // Connect to ERP db
        private SqlConnection connectDb()
        {
            ConfigurationsDB configsDb = ReadAccessDb();

            try
            {
                UpdateLabel("Connecting to Sage DB...", Color.Black, 3);

                var dataSource = configsDb.ServerDb;
                var database = configsDb.Database;
                var userDb = configsDb.UserDb;
                var passwordDb = configsDb.PasswordDb;

                var connectionStringBuilder = new SqlConnectionStringBuilder
                {
                    DataSource = dataSource,
                    InitialCatalog = database,
                    UserID = userDb,
                    Password = passwordDb,
                    PersistSecurityInfo = true
                };

                SqlConnection connection = new SqlConnection(connectionStringBuilder.ConnectionString);
                connection.Open();

                UpdateLabel("Connected to Sage DB.", Color.Black, 3);

                return connection;
            }
            catch (Exception ex)
            {
                UpdateLabel("Error connecting to Sage DB", Color.Red, 3);
                return null;
            }
        }

        private void LogError(string errorMessage)
        {
            // Log file path
            string logFilePath = "error.log";

            try
            {
                // Write the error message to the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"[{DateTime.Now}] {errorMessage}");
                }
            }
            catch (Exception ex)
            {
                // If there is an error writing to the log file, you can log the error elsewhere, such as to the console
                Console.WriteLine("Error writing to log file: " + ex.Message);
            }
        }

        // Make feed
        private bool doFeed(SqlConnection connection)
        {
            // Show time and date of update
            DateTime now = DateTime.Now;
            string dateTimeString = now.ToString("yyyy-MM-dd HH:mm:ss");

            // Insert profit field in the xml file and save it
            try
            {
                updateXmlProfit(connection);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        // Populate the textboxes
        private bool populateInfo()
        {
            // Read config file
            JObject configs = JsonReader.ReadJsonFromFile(configFilePath, toolStripStatusLabel1);

            // Check if there is data in the config file
            if (configs != null)
            {
                // Extract data
                string urlFeedRedicom = (string)configs["urlFeedRedicom"];
                string urlServidorFtp = (string)configs["urlServidorFtp"];
                string utilizador = (string)configs["utilizador"];
                string password = (string)configs["password"];

                // Populate text boxes
                textBox1.Text = urlFeedRedicom;
                textBox2.Text = urlServidorFtp;
                textBox3.Text = utilizador;
                textBox4.Text = password;

                UpdateLabel("Configurations loaded.", Color.Black, 1);

                return true;
            }
            else
            {
                UpdateLabel("Error reading configurations.", Color.Red, 1);
                return false;
            }
        }

        // Read configs for connecting db
        private ConfigurationsDB ReadAccessDb()
        {
            // Read config file
            JObject configs = JsonReader.ReadJsonFromFile(dbConfigFilePath, toolStripStatusLabel1);

            // Check if there is data in the config file
            if (configs != null)
            {
                // Extract data
                ConfigurationsDB configurationsDB = new ConfigurationsDB
                {
                    ServerDb = (string)configs["datasource"],
                    Database = (string)configs["database"],
                    UserDb = (string)configs["utilizadorDb"],
                    PasswordDb = (string)configs["passwordDb"]
                };

                UpdateLabel("DB configurations loaded.", Color.Black, 2);
                return configurationsDB;
            }
            else
            {
                UpdateLabel("Error reading DB configurations.", Color.Red, 2);
                return null;
            }
        }

        // Upload new file to FTP server
        private void UploadFileToFtp(string ftpServer, string ftpUsername, string ftpPassword, string localFilePath, string remoteDirectory)
        {
            // Connect to the FTP server
            FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create($"{ftpServer}{remoteDirectory}/{Path.GetFileName(localFilePath)}");
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
            ftpRequest.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

            // Read the local file
            using (FileStream fileStream = File.OpenRead(localFilePath))
            using (Stream ftpStream = ftpRequest.GetRequestStream())
            {
                // Upload the file to the FTP server
                fileStream.CopyTo(ftpStream);
            }

            // Close the connection
            ftpRequest = null;
        }

        // Insert the profit field and create a new XML file
        private void updateXmlProfit(SqlConnection connection)
        {
            string url = textBox1.Text;

            // Download XML file
            string xmlContent;
            using (WebClient client = new WebClient())
            {
                xmlContent = client.DownloadString(url);
            }

            // Load file
            XDocument doc = XDocument.Parse(xmlContent);

            // Namespace for element 'g:profit'
            XNamespace g = "http://base.google.com/ns/1.0";

            // Update status
            UpdateLabel("Updating feed...", Color.Black, 4);

            // Set the maximum of the ProgressBar based on the number of entries in the XML
            Invoke(new MethodInvoker(() =>
            {
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = doc.Descendants("{http://www.w3.org/2005/Atom}entry").Count();
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = true;
            }));

            int progress = 0;

            // Add the element <g:profit>
            foreach (XElement entry in doc.Descendants("{http://www.w3.org/2005/Atom}entry"))
            {
                // Update progress bar
                Invoke(new MethodInvoker(() =>
                {
                    progress++;
                    toolStripProgressBar1.Value = progress;
                }));

                // Get the item id to query the database
                string itemId = entry.Element(g + "id").Value;

                // Get the actual PVP
                string itemPvp = "";
                itemPvp = entry.Element(g + "sale_price").Value;

                if (string.IsNullOrEmpty(itemPvp))
                {
                    itemPvp = entry.Element(g + "price").Value;
                }

                string pvpFinalString = itemPvp.Replace(" EUR", "").Trim(); // Remove " EUR" and whitespace

                // Get the tax rate
                double taxRate = 0.00;
                int tax = GetTaxRate(connection, itemId);

                if (tax == 1)
                {
                    taxRate = 1.23;
                }
                else if (tax == 3)
                {
                    taxRate = 0.06;
                }
                else
                {
                    taxRate = 1.23;
                }

                // Get the cost price from the db
                string costPrice = GetCostPrice(connection, itemId);

                // Convert prices to float
                float costPriceFloat = float.Parse(costPrice, CultureInfo.InvariantCulture);
                float pvpPriceFloat = float.Parse(pvpFinalString, CultureInfo.InvariantCulture);

                // Calculate margin
                float margin = (float)(pvpPriceFloat - (costPriceFloat + (pvpPriceFloat - (pvpPriceFloat / taxRate))));

                // Convert the margin to string
                string marginValue = margin.ToString("0.00", CultureInfo.InvariantCulture);

                // Define the value of the field <g:profit> 
                entry.Add(new XElement(g + "profit", marginValue));
            }

            // Update progress bar
            Invoke(new MethodInvoker(() =>
            {
                progress++;
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = false;
            }));

            // Update status
            DateTime dateTimeNow = DateTime.Now;
            string dateTimeNowString = dateTimeNow.ToString("dd/MM/yyyy HH:mm:ss");
            UpdateLabel("Feed successfully updated: " + dateTimeNowString + " - Products: " + progress, Color.Black, 4);

            // Save XML
            doc.Save("profit.xml");
        }

        // Get the cost price
        private string GetCostPrice(SqlConnection connection, string itemId)
        {
            // SQL query to get the cost price without VAT of the specific item
            string sqlQuery = "SELECT * FROM dbo.ItemSellingPrices WHERE ItemID = @ItemId AND PriceLineID = 0";

            SqlCommand commandProduct = new SqlCommand(sqlQuery, connection);

            commandProduct.Parameters.AddWithValue("@ItemId", itemId);

            using (SqlDataReader dataReaderProduct = commandProduct.ExecuteReader())
            {
                // Check if rows are returned
                if (dataReaderProduct.HasRows)
                {
                    if (dataReaderProduct.Read())
                    {
                        // Get the column value
                        double rowData = dataReaderProduct.GetDouble(5);
                        // Format the value with two decimal places
                        string formattedValue = rowData.ToString("0.00", CultureInfo.InvariantCulture);
                        return formattedValue;
                    }
                }
            }

            throw new Exception("Error getting the cost price without VAT of the item.");
        }

        // Get the tax rate
        private int GetTaxRate(SqlConnection connection, string itemId)
        {
            // SQL query to get the tax rate of the specific item
            string sqlQuery = "SELECT TaxableGroupID FROM dbo.Item WHERE ItemID = @ItemID";

            SqlCommand commandProduct = new SqlCommand(sqlQuery, connection);

            commandProduct.Parameters.AddWithValue("@ItemId", itemId);

            using (SqlDataReader dataReaderProduct = commandProduct.ExecuteReader())
            {
                // Check if rows are returned
                if (dataReaderProduct.HasRows)
                {
                    if (dataReaderProduct.Read())
                    {
                        // Get the column value
                        int rowData = dataReaderProduct.GetInt32(0);
                        return rowData;
                    }
                }
            }

            throw new Exception("Error getting the tax rate of the item.");
        }

        // Method to update the content of the status strip
        private void UpdateStatus(string statusLabel, string progressLabel, int progressValue)
        {
            toolStripStatusLabel1.Text = statusLabel;
            toolStripStatusLabel2.Text = progressLabel;
            toolStripProgressBar1.Value = progressValue;
        }

        // Method to update the content of a specific label
        private void UpdateLabel(string text, Color textColor, int labelIndex)
        {
            // Check if InvokeRequired is necessary
            if (InvokeRequired)
            {
                // If we are in a different thread than the UI thread, call the method again in the UI thread
                BeginInvoke(new Action(() => UpdateLabel(text, textColor, labelIndex)));
                return;
            }

            // Now we are in the UI thread, we can update the UI controls directly
            switch (labelIndex)
            {
                case 1:
                    toolStripStatusLabel1.Text = text;
                    toolStripStatusLabel1.ForeColor = textColor;
                    break;
                case 2:
                    toolStripStatusLabel2.Text = text;
                    toolStripStatusLabel2.ForeColor = textColor;
                    break;
                case 3:
                    toolStripStatusLabel3.Text = text;
                    toolStripStatusLabel3.ForeColor = textColor;
                    break;
                case 4:
                    toolStripStatusLabel4.Text = text;
                    toolStripStatusLabel4.ForeColor = textColor;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("labelIndex", "Label index should be between 1 and 4.");
            }
        }

        #endregion

        #region ButtonClicks

        private void editBtn_Click(object sender, EventArgs e)
        {
            editBtn.Enabled = false;
            button1.Enabled = false;
            clrBtn.Enabled = true;
            saveBtn.Enabled = true;
            readConfigBtn.Enabled = true;
            cancelBtn.Enabled = true;
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
        }

        private void clrBtn_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            // Initial state of buttons
            initStat();

            // Update fields
            populateInfo();
        }

        private void readConfigBtn_Click(object sender, EventArgs e)
        {
            // Update fields
            populateInfo();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            bool fieldsNotEmpty = true;

            // Example of calling WriteJsonToFile method
            string urlFeedRedicom = textBox1.Text; // Get user input from TextBox1
            string urlServidorFtp = textBox2.Text; // Get user input from TextBox2
            string utilizador = textBox3.Text; // Get user input from TextBox3
            string password = textBox4.Text; // Get user input from TextBox4

            // Check fields
            // Check if all fields are filled
            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox3.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                // If any field is empty, set the variable to false and show the error message
                fieldsNotEmpty = false;
                MessageBox.Show("All fields must be filled", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (fieldsNotEmpty)
            {
                // Call the WriteJsonToFile method
                JsonWriter.WriteJsonToFile(configFilePath, urlFeedRedicom, urlServidorFtp, utilizador, password);

                // Update fields
                populateInfo();

                // Initial state of buttons
                initStat();
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // Initial state of components
            initStat();

            await LoadDataAsync();
        }

        #endregion
    }

    #region Classes

    // Read the config file
    class JsonReader
    {
        public static JObject ReadJsonFromFile(string filePath, ToolStripStatusLabel toolStripStatusLabel1)
        {
            try
            {
                // File exists?
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("Configuration file not found.", filePath);
                }

                // Read content
                string jsonString = File.ReadAllText(filePath);

                // Parse to JSON
                JObject json = JObject.Parse(jsonString);

                return json;
            }
            catch (FileNotFoundException ex)
            {
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }

    // Write to config file
    class JsonWriter
    {
        public static void WriteJsonToFile(string filePath, string urlFeedRedicom, string urlServidorFtp, string utilizador, string password)
        {
            try
            {
                // Create JSON object
                JObject json = new JObject(
                    new JProperty("urlFeedRedicom", urlFeedRedicom),
                    new JProperty("urlServidorFtp", urlServidorFtp),
                    new JProperty("utilizador", utilizador),
                    new JProperty("password", password)
                );

                // Write JSON to file
                File.WriteAllText(filePath, json.ToString());

                // Success message
                MessageBox.Show("Data saved successfully.");
            }
            catch (Exception ex)
            {
                // Error message
                MessageBox.Show("Could not save the data:" + ex.Message);
            }
        }
    }


    class ConfigurationsDB
    {
        public string ServerDb { get; set; }
        public string Database { get; set; }
        public string UserDb { get; set; }
        public string PasswordDb { get; set; }
    }

    #endregion

}
