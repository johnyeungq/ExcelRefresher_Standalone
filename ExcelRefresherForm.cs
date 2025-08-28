using JYLIB;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using _Excel = Microsoft.Office.Interop.Excel;   

namespace ExcelRefresher_Standalone
{
    public partial class ExcelRefresherForm : Form
    {
        JYLIB.Main _main;

        DataClasses _data;

        bool IsAuto { get; set; }

        public ExcelRefresherForm(bool Auto)
        {

            InitializeComponent();
            if (_main == null)
            {
                _main = new JYLIB.Main();
            }
            if (_data == null)
            {
                _data = new DataClasses();
            }
            IsAuto = Auto;

        }


        public async Task RefreshExcel(string file)
        {
            
            //_main.Log("Function Dropped successfully");
            _main.KillExcel();
            _Excel.Application excelApp = new _Excel.Application();
            _Excel.Workbook workbook = excelApp.Workbooks.Open(file);
            excelApp.DisplayAlerts = false;

            workbook.RefreshAll();
            await Task.Delay(500);

            bool saveSuccessful = false;
            int maxRetryAttempts = 3;
            int retryCount = 0;

            while (!saveSuccessful && retryCount < maxRetryAttempts)
            {
                try
                {
                    workbook.Save();
                    saveSuccessful = true;
                    Console.WriteLine(file + " Refreshed!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
                    retryCount++;
                    await Task.Delay(5 * 1000);
                }
            }

            retryCount = 0;
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            string Gnow = DateTime.Now.ToString("G");
            _main.KillExcel();
            await Task.Delay(2 * 1000);


            _main.Log(Gnow + " Task - " + file + " is refreshed" + "\n");

            await Task.Delay(2 * 1000);

        }
        public List<string> XMLfile(string xmlFilePath, string child, string Node)
        {
            List<string> var = new List<string>();
   
        // Load the XML file
           XmlDocument xmlDoc = new XmlDocument();
   
            xmlDoc.Load(xmlFilePath);

            // Get the root element of the XML document
            XmlElement rootElement = xmlDoc.DocumentElement;

            // Loop through eac;h child node of the root element
            foreach (XmlNode childNode in rootElement.ChildNodes)
            {
                // Check if the child node is a "Locations" node
                if (childNode.Name == child)
                {
                    // Get the value of the "Name" element
                    XmlNode nameNode = childNode.SelectSingleNode(Node);
                    if (nameNode != null)
                    {
                        string str = nameNode.InnerText;
                        var.Add(str);
                    }
                }
            }

            return var;
        }
        private async void ExcelRefresherForm_Load(object sender, EventArgs e)
        {
            XMLpathTB.Text = _data.XML;
          

                ConfigPathTB.Text = _data.Config;

          //  logfolderPathTB.Text = _data.FetchLogPath();
            if (IsAuto)
            {
                tabControl1.Controls.Remove(ManualPage);

            }
            else
            {
                tabControl1.Controls.Remove(AutoPage);
            }



            List<string> excelpath = XMLfile(_data.XML, "Excel", "ExcelPath");
            List<string> excelName = XMLfile(_data.XML, "Excel", "ExcelName");
            foreach (string excel in excelpath)
            {
                ExcelPathLB.Items.Add(excel);
                AutoLB.Items.Add(excel);
            }



            foreach (string excel in excelName)
            {
                ExcelNameLB.Items.Add(excel);
            }



            if (IsAuto)
            {
                ExcelCountLB.Text = $"Excel Count : {AutoLB.Items.Count.ToString()} ";
                count = AutoLB.Items.Count;
                await AutoRefresh();


            }



        }
        void AutoLog(string log)
        {

            AutoLogLB.Items.Add(log);
            //string today = DateTime.Now.ToString("yyyy-MM-dd");
            _main.Log(log);
          

        }


        int count = 0;
        int successfulRefreshes = 0;
        async Task AutoRefresh()
        {


            progressBar1.Maximum = AutoLB.Items.Count;
            progressBar1.Value = 0;

            if (AutoLB.Items.Count < 1)
            {
                MessageBox.Show("No Excel files found to refresh.");
                return;
            }
            else
            {
                //  _main.KillExcel();
                foreach (string excel in AutoLB.Items)
                {



                    try
                    {
                        // await _main.RefreshExcel(excel);
                        await RefreshExcel(excel);
                        AutoLog(excel + " refreshing at " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        await Task.Delay(30 * 1000);
                        progressBar1.Value++;
                        successfulRefreshes++;
                        AutoLog(excel + " refreshed at " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error refreshing {excel}: {ex.Message}");
                        continue;
                    }
                    string donetempFolder = @"C:\data\donetemp\";
                    if(!Directory.Exists(donetempFolder))
                    {
                        Directory.CreateDirectory(donetempFolder);
                    }

                    if (successfulRefreshes == AutoLB.Items.Count)
                    {
                        // MessageBox.Show("All Excel files refreshed successfully!");
                        string filename = $@"{donetempFolder}\{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.Done";
                        string filename2 = Path.Combine (donetempFolder, DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")+".Done");        
                        File.Create(filename2).Close();
                        //MessageBox.Show(filename);
                        Application.Exit();
                    }
                    else if (successfulRefreshes > 0)
                    {
                        //   MessageBox.Show($"{successfulRefreshes} of {AutoLB.Items.Count} Excel files refreshed successfully.");
                    }
                    else
                    {
                        //   MessageBox.Show("No Excel files were refreshed successfully.");
                    }
                    await Task.Delay(30 * 1000);
                }
            }




        }

        private async void RefreshExcelBtn_Click(object sender, EventArgs e)
        {
            _main.KillExcel();
            if (ExcelPathLB.SelectedIndex < 0)
            {
                MessageBox.Show("Please select an Excel file to refresh.");
                return;
            }
            else
            {

              _main.RefreshExcel(ExcelPathLB.SelectedItem.ToString());
                // MessageBox.Show("");


            }



        }

        private void ExcelNameLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ExcelNameLB.SelectedIndex < 0)
            {
                return;
            }
            else
            {
                ExcelPathLB.SelectedIndex = ExcelNameLB.SelectedIndex;
            }
        }

        private void ExcelPathLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ExcelPathLB.SelectedIndex < 0)
            {
                return;
            }
            else
            {
                ExcelNameLB.SelectedIndex = ExcelPathLB.SelectedIndex;
            }
        }

        private void AddExcelBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(AddexcelPathTb.Text) || string.IsNullOrEmpty(AddexcelNameTB.Text))
            {
                MessageBox.Show("Please enter both Excel Path and Name.");
                return;
            }
            else {

                AddFilePathToXml(_data.XML, AddexcelNameTB.Text, AddexcelPathTb.Text, _data.Appfolder + @"\ExcelFiles\"); 
            }
        }
        internal void AddFilePathToXml(string xmlFilePath, string excelName, string excelPath, string destinationFolder = null)
        {
            try
            {
                if (string.IsNullOrEmpty(xmlFilePath))
                {
                    _main.Log("XML file path is empty or null.");
                    throw new ArgumentException("XML file path cannot be empty or null.");
                }

                if (string.IsNullOrEmpty(excelName) || string.IsNullOrEmpty(excelPath))
                {
                    _main.Log("Excel file name or path is empty or null.");
                    throw new ArgumentException("Excel file name or path cannot be empty or null.");
                }

                XDocument doc;
                if (File.Exists(xmlFilePath))
                {
                    doc = XDocument.Load(xmlFilePath);
                }
                else
                {
                    doc = new XDocument(
                        new XDeclaration("1.0", "UTF-8", null),
                        new XElement("dataroot", new XAttribute(XNamespace.Xmlns + "od", "urn:schemas-microsoft-com:officedata"), new XAttribute("generated", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")))
                    );
                    _main.Log($"Created new XML file at '{xmlFilePath}'.");
                }

                var existingPaths = doc.Descendants("Excel")
                    .Elements("ExcelPath")
                    .Select(e => e.Value)
                    .ToList();

                if (existingPaths.Contains(excelPath, StringComparer.OrdinalIgnoreCase))
                {
                    _main.Log($"File path '{excelPath}' already exists in '{xmlFilePath}'.");
                    return;
                }

                doc.Element("dataroot").Add(
                    new XElement("Excel",
                        new XElement("ExcelName", excelName),
                        new XElement("ExcelPath", excelPath)
                    )
                );

                doc.Save(xmlFilePath);
                _main.Log($"Added '{excelName}' with path '{excelPath}' to '{xmlFilePath}'.");

                // Optionally copy the file to destinationFolder
                if (!string.IsNullOrEmpty(destinationFolder) && Directory.Exists(destinationFolder))
                {
                    if (File.Exists(excelPath))
                    {
                        string destPath = Path.Combine(destinationFolder, excelName);
                        File.Copy(excelPath, destPath, true);
                        _main.Log($"Copied '{excelPath}' to '{destPath}'.");
                    }
                    else
                    {
                        _main.Log($"File '{excelPath}' does not exist and was not copied.");
                    }
                 

                }
            }
            catch (Exception ex)
            {
                _main.Log($"Error adding file path to '{xmlFilePath}': {ex.Message}");
                throw;
            }

            refresh();  
        }

        void refresh() {

            ExcelNameLB.Items.Clear();
            AutoLB.Items.Clear();
            ExcelPathLB.Items.Clear();
            List<string> excelpath2 = XMLfile(_data.XML, "Excel", "ExcelPath");
            List<string> excelName2 = XMLfile(_data.XML, "Excel", "ExcelName");
            foreach (string excel in excelpath2)
            {
                ExcelPathLB.Items.Add(excel);
                AutoLB.Items.Add(excel);
            }



            foreach (string excel in excelName2)
            {
                ExcelNameLB.Items.Add(excel);
            }
        }
        private void AutoPage_Click(object sender, EventArgs e)
        {//

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
