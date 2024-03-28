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

            //Initial state of components
            initStat();

            // Configurar o temporizador
            hourlyTimer = new System.Windows.Forms.Timer();
            hourlyTimer.Interval = 3600000; // 1 hora em milissegundos
            hourlyTimer.Tick += HourlyTimer_Tick;

            //First call to update feed
            await LoadDataAsync();

            // Iniciar o temporizador para chamar a função a cada hora
            hourlyTimer.Start();

        }

        #region Trayicon

        //Minimize to tray icon
        private void Form1_Resize(object sender, EventArgs e)
        {
            //if the form is minimized
            //hide it from the task bar
            //and show the system tray icon (represented by the NotifyIcon control)
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }

        //Show form again on tray icon click
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
            // Chamar a função LoadDataAsync
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

                    UpdateLabel("Desligado da BD do Sage.", Color.Black, 3);
                }
            }

            button1.Enabled = true;
            editBtn.Enabled = true;
        }

        //Initial state
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

        //Connect to ERP db
        private SqlConnection connectDb()
        {
            ConfigurationsDB configsDb = ReadAccessDb();

            try {

                UpdateLabel("A ligar à BD do Sage...", Color.Black, 3);

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

                UpdateLabel("Ligado à BD do Sage.", Color.Black, 3);

                return connection;

            }
            catch (Exception ex)
            {
                UpdateLabel("Erro a ligar à BD do Sage", Color.Red, 3);

                return null;
            }

        }

        private void LogError(string errorMessage)
        {
            // Caminho do arquivo de log
            string logFilePath = "error.log";

            try
            {
                // Grava a mensagem de erro no arquivo de log
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"[{DateTime.Now}] {errorMessage}");

                }
            }
            catch (Exception ex)
            {
                // Se ocorrer algum erro ao gravar no arquivo de log, você pode registrar o erro em outro lugar, como no console
                Console.WriteLine("Erro ao gravar no arquivo de log: " + ex.Message);
            }
        }

        //Make feed
        private bool doFeed(SqlConnection connection)
        {

            //Show time and date of update
            DateTime now = DateTime.Now;
            string dateTimeString = now.ToString("yyyy-MM-dd HH:mm:ss");

            //Insert profit field in the xml file and saves it
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

        //Populate the textboxes
        private bool populateInfo()
        {
            // Read config file
            JObject configs = JsonReader.ReadJsonFromFile(configFilePath, toolStripStatusLabel1);

            // There is data in config file
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

                UpdateLabel("Configurações carregadas.", Color.Black, 1);

                return true;
            }
            else {
                UpdateLabel("Erro a ler configurações.", Color.Red, 1);
                return false; 
            }
        }

        //Read configs for connect db
        private ConfigurationsDB ReadAccessDb()
        {
            // Read config file
            JObject configs = JsonReader.ReadJsonFromFile(dbConfigFilePath, toolStripStatusLabel1);

            // There is data in config file
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

                UpdateLabel("Configurações DB carregadas.", Color.Black, 2);
                return configurationsDB;

            }
            else {

                UpdateLabel("Erro a ler configurações DB.", Color.Red, 2);
                return null;
            
            }
        }

        //Upload new file to ftp server
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

        //Insert the profit field and create new xml file
        private void updateXmlProfit(SqlConnection connection)
        {
            string url = textBox1.Text;

            // Download xml file
            string xmlContent;
            using (WebClient client = new WebClient())
            {
                xmlContent = client.DownloadString(url);
            }

            // Load file
            XDocument doc = XDocument.Parse(xmlContent);

            // Namespace for element 'g:profit'
            XNamespace g = "http://base.google.com/ns/1.0";

            //Atualiza status
            UpdateLabel("A fazer o feed...", Color.Black, 4);

            // Define o máximo da ProgressBar com base no número de entradas no XML
            Invoke(new MethodInvoker(() =>
            {
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = doc.Descendants("{http://www.w3.org/2005/Atom}entry").Count();
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = true;
            }));
           
            int progress = 0;
            //
            //toolStripStatusLabel4.Text = "A fazer o feed...";

            // Add the element <g:profit>
            foreach (XElement entry in doc.Descendants("{http://www.w3.org/2005/Atom}entry"))
            {
                // Atualiza a barra de progresso
                Invoke(new MethodInvoker(() =>
                {
                    progress++;
                    toolStripProgressBar1.Value = progress;
                }));

                //Get the item id to query database
                string itemId = entry.Element(g + "id").Value;

                //Get the actual PVP
                string itemPvp = "";
                itemPvp = entry.Element(g + "sale_price").Value;
                
                if (string.IsNullOrEmpty(itemPvp))
                {
                    itemPvp = entry.Element(g + "price").Value;
                }
                else
                {
                    ;
                }

                string pvpFinalString = itemPvp.Replace(" EUR", "").Trim(); // Remover " EUR" e espaços em branco

                //Get the tax rate
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
                //Debug
                //LogError(itemId + ": " + taxRate.ToString(CultureInfo.InvariantCulture) + "\n");

                //Get the cost price from db
                string costPrice = GetCostPrice(connection, itemId);

                //Convert prices  to float
                float costPriceFloat = float.Parse(costPrice, CultureInfo.InvariantCulture);
                float pvpPriceFloat = float.Parse(pvpFinalString, CultureInfo.InvariantCulture);

                //Calculate margin
                float margin = (float)(pvpPriceFloat - (costPriceFloat + (pvpPriceFloat - (pvpPriceFloat / taxRate))));

                //Convert the margin to string
                string marginValue = margin.ToString("0.00", CultureInfo.InvariantCulture);

                // Define the value of the field <g:profit> 
                entry.Add(new XElement(g + "profit", marginValue));

            }

            // Atualiza a barra de progresso
            Invoke(new MethodInvoker(() =>
            {
                progress++;
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Visible = false;

            }));

            //Atualiza status
            DateTime dateTimeNow = DateTime.Now;
            string dateTimeNowString = dateTimeNow.ToString("dd/MM/yyyy HH:mm:ss");
            UpdateLabel("Feed efectuado com sucesso: " + dateTimeNowString + " -Produtos: " + progress, Color.Black, 4);

            // Save XML
            doc.Save("profit.xml");
        }

        // Get the cost price
        private string GetCostPrice(SqlConnection connection, string itemId)
        {
            // Consulta SQL para obter o preço de custo sem IVA do item específico
            string sqlQuery = "SELECT * FROM dbo.ItemSellingPrices WHERE ItemID = @ItemId AND PriceLineID = 0";

            SqlCommand commandProduct = new SqlCommand(sqlQuery, connection);
            
            commandProduct.Parameters.AddWithValue("@ItemId", itemId);

            using (SqlDataReader dataReaderProduct = commandProduct.ExecuteReader())
            {

                // Verificar se há linhas retornadas
                if (dataReaderProduct.HasRows)
                {
                    if (dataReaderProduct.Read())
                    {
                       
                        // Obtenha o valor da coluna
                        double rowData = dataReaderProduct.GetDouble(5);
                        // Debug.WriteLine("Valor da coluna: " + rowData);

                        // Formate o valor com duas casas decimais
                        string formattedValue = rowData.ToString("0.00", CultureInfo.InvariantCulture);
                        // Debug.WriteLine("Valor formatado: " + formattedValue);

                        return formattedValue;
                    }
                }

            }

            throw new Exception("Erro ao obter o preço de custo sem IVA do item.");
     
        }

        // Get the cost price
        private int GetTaxRate(SqlConnection connection, string itemId)
        {
            // Consulta SQL para obter o preço de custo sem IVA do item específico
            string sqlQuery = "SELECT TaxableGroupID FROM dbo.Item WHERE ItemID = @ItemID";

            SqlCommand commandProduct = new SqlCommand(sqlQuery, connection);

            commandProduct.Parameters.AddWithValue("@ItemId", itemId);
            
            using (SqlDataReader dataReaderProduct = commandProduct.ExecuteReader())
            {

                // Verificar se há linhas retornadas
                if (dataReaderProduct.HasRows)
                {
                    if (dataReaderProduct.Read())
                    {

                        // Obtenha o valor da coluna
                        int rowData = dataReaderProduct.GetInt32(0);
                        
                        // Formate o valor com duas casas decimais
                        //string formattedValue = rowData.ToString("0.00", CultureInfo.InvariantCulture);

                        return rowData;
                    }
                }

            }

            throw new Exception("Erro ao obter o preço de custo sem IVA do item.");

        }


        // Método para atualizar o conteúdo da status strip
        private void UpdateStatus(string statusLabel, string progressLabel, int progressValue)
        {
            toolStripStatusLabel1.Text = statusLabel;
            toolStripStatusLabel2.Text = progressLabel;
            toolStripProgressBar1.Value = progressValue;
        }

        // Método para atualizar o conteúdo de uma label específica
        // Método para atualizar o conteúdo de uma label específica com texto e cor
        private void UpdateLabel(string text, Color textColor, int labelIndex)
        {
            // Verifica se é necessário chamar InvokeRequired
            if (InvokeRequired)
            {
                // Se estamos em uma thread diferente da thread de IU, chamamos novamente o método na thread de IU
                BeginInvoke(new Action(() => UpdateLabel(text, textColor, labelIndex)));
                return;
            }

            // Agora estamos na thread de IU, podemos atualizar os controles da interface do usuário diretamente
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
                    throw new ArgumentOutOfRangeException("labelIndex", "Índice da label deve estar entre 1 e 4.");
            }
        }



        #endregion

        #region ButonClicks

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
       
            // Init state of buttons
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

            //Check fields
            // Verifica se todos os campos estão preenchidos
            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox3.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                // Se algum campo estiver vazio, define a variável como false e exibe a mensagem de erro
                fieldsNotEmpty = false;
                MessageBox.Show("Todos os campos devem ser preenchidos", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (fieldsNotEmpty)
            {

                // Call the WriteJsonToFile method
                JsonWriter.WriteJsonToFile(configFilePath, urlFeedRedicom, urlServidorFtp, utilizador, password);

                // Update fields
                populateInfo();

                // Init state of buttons
                initStat(); 

            }

            

        }
       

        private async void button1_Click(object sender,  EventArgs e)
        {
            //Initial state of components
            initStat();

            await LoadDataAsync();

        }
        
        #endregion
    }

    #region Classes

    //Read the config file
    class JsonReader
    {
        public static JObject ReadJsonFromFile(string filePath, ToolStripStatusLabel toolStripStatusLabel1)
        {
            try
            {
                // file exists?
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("Ficheiro de configuração não encontrado.", filePath);
                }

                // Read content
                string jsonString = File.ReadAllText(filePath);

                // Parse to JSON
                JObject json = JObject.Parse(jsonString);

                return json;
            }
            catch (FileNotFoundException ex)
            {
                // Atualiza o texto e a cor do ToolStripStatusLabel para indicar o erro
                //toolStripStatusLabel1.Text = "Ficheiro de configuração não encontrado";
                //toolStripStatusLabel1.ForeColor = Color.Red;
                return null;
            }
            catch (Exception ex)
            {
                //toolStripStatusLabel1.Text = "Erro ao ler o ficheiro de configuração:" + ex.Message;
                //toolStripStatusLabel1.ForeColor = Color.Red;
                return null;
            }
        }
    }

    //Write to config file
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

                //Success message
                MessageBox.Show("Dados guardados com sucesso.");
            }
            catch (Exception ex)
            {
                // Error message
                MessageBox.Show("Não foi possível guardar os dados:" + ex.Message);
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
