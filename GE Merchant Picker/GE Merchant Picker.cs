using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Microsoft.Office.Interop.Excel;
using System.IO;
using AutoIt;
using System.Data.SqlClient;
using System.Net;
using System.Drawing;
using System.Text;

namespace GE_Merchant_Picker
{
    enum EnvironmentType
    {
        QA,
        Staging,
        Production
    }

    class DAL
    {
        static public String readFromSQL(String query, String columnName, string connectionString)
        {
            string SQLResult = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection())
                {
                    connection.ConnectionString = connectionString;
                    connection.Open();
                    using (SqlCommand myCommand = new SqlCommand(query, connection))
                    using (SqlDataReader myReader = myCommand.ExecuteReader())
                    {
                        while (myReader.Read())
                        {
                            SQLResult = myReader[columnName].ToString();
                        }
                    }
                }
            }
            catch (Exception) 
            {
                MessageBox.Show("Can't take read data from DB");
            }

            return SQLResult;

        }

    }
    class EnvironmentData
    {
        public int startRowInExcelWithMerchantName;
        public List<String> merchantsList { get; private set; } = new List<string>();
        public Worksheet xlWorksheet;
        public Range xlRange;
        private string ConnectionString;

        public EnvironmentData(int startRowInExcelWithMerchantName, Worksheet xlWorksheet, string serverIP = null)
        {
            this.startRowInExcelWithMerchantName = startRowInExcelWithMerchantName;
            this.xlWorksheet = xlWorksheet;
            this.xlRange = xlWorksheet.UsedRange;

            int tempRow = startRowInExcelWithMerchantName;


            while (xlRange.Cells[tempRow, 1] != null && xlRange.Cells[tempRow, 1].Value2 != null)
            {
                merchantsList.Add(xlRange.Cells[tempRow, 1].Value2.ToString());
                tempRow++;
            }

            if (!string.IsNullOrWhiteSpace(serverIP))
            {
                ConnectionString = "user id=sql_qa_ukr_DenisH;" +
                    "password=Admin_141;" +
                    "server=" +
                    serverIP +
                    ";" +
                    "Trusted_Connection=no;" +
                    "database=GlobalE;" +
                    "connection timeout=30";
            }

        }

        public Merchant GetMerchant(string merchantName)
        {
            int merchantRow = merchantsList.IndexOf(merchantName) + startRowInExcelWithMerchantName;

            var merchant = new Merchant();

            merchant.merchantName = xlRange.Cells[merchantRow, 1].Value2?.ToString() ?? String.Empty;

            merchant.merchantSiteUri = xlRange.Cells[merchantRow, 2].Value2?.ToString() ?? String.Empty;
            merchant.adminUri = xlRange.Cells[merchantRow, 3].Value2?.ToString() ?? String.Empty;
            merchant.adminLoginUserName = xlRange.Cells[merchantRow, 4].Value2?.ToString() ?? String.Empty;
            merchant.adminLoginPassword = xlRange.Cells[merchantRow, 5].Value2?.ToString() ?? String.Empty;
            merchant.mid = xlRange.Cells[merchantRow, 6].Value2?.ToString() ?? String.Empty;
            merchant.siteLoginUserName = xlRange.Cells[merchantRow, 7].Value2?.ToString() ?? String.Empty;
            merchant.siteLoginPassword = xlRange.Cells[merchantRow, 8].Value2?.ToString() ?? String.Empty;

            merchant.comments = xlRange.Cells[merchantRow, 9].Value2?.ToString() ?? String.Empty;
            merchant.returnPortalUri = xlRange.Cells[merchantRow, 10].Value2?.ToString() ?? String.Empty;
            merchant.logsUri = xlRange.Cells[merchantRow, 11].Value2?.ToString() ?? String.Empty;
            merchant.coupons = xlRange.Cells[merchantRow, 12].Value2?.ToString() ?? String.Empty;
            merchant.trackingPortalUri = xlRange.Cells[merchantRow, 13].Value2?.ToString() ?? String.Empty;

            if (merchant.mid == "")
            {

                String queryMid = "select top 1 MerchantId from Merchants where merchantname like '%" + merchant.merchantName + "%'"
                                    + " and IsActive = 1 and SiteURL = '" + merchant.merchantSiteUri + "'";
                merchant.mid = DAL.readFromSQL(queryMid, "MerchantId", ConnectionString);
            }

            String queryPlatform = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (select top 1 MerchantPlatformId from Merchants where merchantname like '%"
                                + merchant.merchantName + "%' and IsActive = 1)";
            merchant.platformType = DAL.readFromSQL(queryPlatform, "MerchantPlatformName", ConnectionString);

            String lineForTextBox = "";
            //StringBuilder 

            if (!string.IsNullOrWhiteSpace(merchant.mid))
            {
                lineForTextBox = lineForTextBox + "MerchantID --> " + merchant.mid;
            }
            if (merchant.platformType != "") lineForTextBox = lineForTextBox + "\nPlatform --> " + merchant.platformType;
            if (merchant.merchantSiteUri != "") lineForTextBox = lineForTextBox + "\nURL -->  " + merchant.merchantSiteUri;
            if (merchant.siteLoginUserName != "") lineForTextBox = lineForTextBox + "\nUser -->  " + merchant.siteLoginUserName;
            if (merchant.siteLoginPassword != "") lineForTextBox = lineForTextBox + "\nPass -->  " + merchant.siteLoginPassword;
            if (merchant.adminUri != "") lineForTextBox = lineForTextBox + "\nAdmin --> " + merchant.adminUri;
            if (merchant.adminLoginUserName != "") lineForTextBox = lineForTextBox + "\nUser -->  " + merchant.adminLoginUserName;
            if (merchant.adminLoginPassword != "") lineForTextBox = lineForTextBox + "\nPass -->  " + merchant.adminLoginPassword;
            if (merchant.returnPortalUri != "") lineForTextBox = lineForTextBox + "\nRetrun Portal --> " + merchant.returnPortalUri;
            if (merchant.trackingPortalUri != "") lineForTextBox = lineForTextBox + "\nTracking Portal --> " + merchant.trackingPortalUri;
            if (merchant.coupons != "") lineForTextBox = lineForTextBox + "\nCoupons --> " + merchant.coupons;
            if (merchant.comments != "") lineForTextBox = lineForTextBox + "\nComment --> " + merchant.comments;


            return merchant;
        }

    }

    public partial class GE_Merchant_Picker_Form : Form
    {
        
        const String GEAdminQA = "https://qa.bglobale.com/GlobaleAdmin";
        const String GEAdminStg = "https://www2.bglobale.com/GlobaleAdmin";
        const String GEAdminProd = "https://web.global-e.com/GlobaleAdmin";

        bool firstRun = true;
        EnvironmentType chosenEnvironment = EnvironmentType.QA;

        String SQLResult = "";
        SqlConnection mySQLConnection = new SqlConnection();

        Merchant selectedMerchant = new Merchant();

        Dictionary<EnvironmentType, EnvironmentData> environmentList = new Dictionary<EnvironmentType, EnvironmentData>();

        //List<String> merchantsListQA = new List<string>();
        //List<String> merchantsListStg = new List<string>();
        //List<String> merchantsListProd = new List<string>();

        static string fileName = "Auto Merchants Adresses.xlsx";
        static string path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\" + fileName);

        //Create COM Objects. Create a COM object for everything that is referenced
        static Microsoft.Office.Interop.Excel.Application xlAppQA = new Microsoft.Office.Interop.Excel.Application();
        static Workbook xlWorkbook = xlAppQA.Workbooks.Open(path);


        public GE_Merchant_Picker_Form()
        {
            environmentList.Add(EnvironmentType.QA, new EnvironmentData(4, xlWorkbook.Sheets["QA"], "54.72.115.215"));
            environmentList.Add(EnvironmentType.Staging, new EnvironmentData(5, xlWorkbook.Sheets["Staging"], "54.72.120.2"));
            environmentList.Add(EnvironmentType.Production, new EnvironmentData(3, xlWorkbook.Sheets["Production"]));

            InitializeComponent();
            initializeMerchantsListBox();

        }

        public void initializeMerchantsListBox()
        {
           merchantsListBox.DataSource = environmentList[chosenEnvironment].merchantsList;
        }

        public void showMerchantDetails(string merchant)
        {

            //int tempRow = environmentList[EnvironmentType.QA].startRowInExcelWithMerchantName;
            //Range xlRange = environmentList[EnvironmentType.QA].xlRange;

            selectedMerchant = environmentList[chosenEnvironment].GetMerchant(merchant);

            //if (chosenEnvironment != EnvironmentType.Production)
            //{
            //    if (chosenEnvironment == EnvironmentType.QA)
            //    {
            //        tempRow = merchantsListQA.IndexOf(merchant) + environmentList[EnvironmentType.QA].startRowInExcelWithMerchantName;
            //        xlRange = environmentList[EnvironmentType.QA].xlRange;
            //    }
            //    else if (chosenEnvironment == EnvironmentType.Staging)
            //    {
            //        tempRow = merchantsListStg.IndexOf(merchant) + environmentList[EnvironmentType.Staging].startRowInExcelWithMerchantName;
            //        xlRange = environmentList[EnvironmentType.Staging].xlRange;
            //    }

            //Initialize selected merchant
            //
            //selectedMerchant.merchantSiteUri = xlRange.Cells[tempRow, 2].Value2 != null ? xlRange.Cells[tempRow, 2].Value2.ToString() : "";

            //selectedMerchant.merchantName = xlRange.Cells[tempRow, 1].Value2?.ToString() ?? String.Empty;

            //selectedMerchant.merchantSiteUri = xlRange.Cells[tempRow, 2].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.adminUri = xlRange.Cells[tempRow, 3].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.adminLoginUserName = xlRange.Cells[tempRow, 4].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.adminLoginPassword = xlRange.Cells[tempRow, 5].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.mid = xlRange.Cells[tempRow, 6].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.siteLoginUserName = xlRange.Cells[tempRow, 7].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.siteLoginPassword = xlRange.Cells[tempRow, 8].Value2?.ToString() ?? String.Empty;

            //selectedMerchant.comments = xlRange.Cells[tempRow, 9].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.returnPortalUri = xlRange.Cells[tempRow, 10].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.logsUri = xlRange.Cells[tempRow, 11].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.coupons = xlRange.Cells[tempRow, 12].Value2?.ToString() ?? String.Empty;
            //selectedMerchant.trackingPortalUri = xlRange.Cells[tempRow, 13].Value2?.ToString() ?? String.Empty;

            //if (selectedMerchant.mid == "")
            //{

            //    String queryMid = "select top 1 MerchantId from Merchants where merchantname like '%" + selectedMerchant.merchantName + "%'"
            //                        + " and IsActive = 1 and SiteURL = '" + selectedMerchant.merchantSiteUri + "'";
            //    selectedMerchant.mid = readFromSQL(queryMid, "MerchantId");
            //}

            //String queryPlatform = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (select top 1 MerchantPlatformId from Merchants where merchantname like '%"
            //                    + selectedMerchant.merchantName + "%' and IsActive = 1)";
            //selectedMerchant.platformType = readFromSQL(queryPlatform, "MerchantPlatformName");

            String lineForTextBox = "";
            //StringBuilder 

            if (!string.IsNullOrWhiteSpace(selectedMerchant.mid))
            {
                lineForTextBox = lineForTextBox + "MerchantID --> " + selectedMerchant.mid;
            }
            if (selectedMerchant.platformType != "") lineForTextBox = lineForTextBox + "\nPlatform --> " + selectedMerchant.platformType;
            if (selectedMerchant.merchantSiteUri != "") lineForTextBox = lineForTextBox + "\nURL -->  " + selectedMerchant.merchantSiteUri;
            if (selectedMerchant.siteLoginUserName != "") lineForTextBox = lineForTextBox + "\nUser -->  " + selectedMerchant.siteLoginUserName;
            if (selectedMerchant.siteLoginPassword != "") lineForTextBox = lineForTextBox + "\nPass -->  " + selectedMerchant.siteLoginPassword;
            if (selectedMerchant.adminUri != "") lineForTextBox = lineForTextBox + "\nAdmin --> " + selectedMerchant.adminUri;
            if (selectedMerchant.adminLoginUserName != "") lineForTextBox = lineForTextBox + "\nUser -->  " + selectedMerchant.adminLoginUserName;
            if (selectedMerchant.adminLoginPassword != "") lineForTextBox = lineForTextBox + "\nPass -->  " + selectedMerchant.adminLoginPassword;
            if (selectedMerchant.returnPortalUri != "") lineForTextBox = lineForTextBox + "\nRetrun Portal --> " + selectedMerchant.returnPortalUri;
            if (selectedMerchant.trackingPortalUri != "") lineForTextBox = lineForTextBox + "\nTracking Portal --> " + selectedMerchant.trackingPortalUri;
            if (selectedMerchant.coupons != "") lineForTextBox = lineForTextBox + "\nCoupons --> " + selectedMerchant.coupons;
            if (selectedMerchant.comments != "") lineForTextBox = lineForTextBox + "\nComment --> " + selectedMerchant.comments;

            richTextBox1.Text = lineForTextBox;

            //}
            //else if (chosenEnvironment == EnvironmentType.Production)
            //{

            //    tempRow = merchantsListProd.IndexOf(merchant) + environmentList[EnvironmentType.Production].startRowInExcelWithMerchantName;
            //    xlRange = environmentList[EnvironmentType.Production].xlRange;

            //    selectedMerchant.merchantName = (xlRange.Cells[tempRow, 1].Value2 ?? String.Empty).ToString();
            //    selectedMerchant.merchantSiteUri = (xlRange.Cells[tempRow, 2].Value2 ?? String.Empty).ToString();
            //    selectedMerchant.mid = (xlRange.Cells[tempRow, 3].Value2 ?? String.Empty).ToString();
            //    selectedMerchant.coupons = (xlRange.Cells[tempRow, 4].Value2 ?? String.Empty).ToString();

                //String lineForTextBox = "";

                //if (selectedMerchant.mid != "") lineForTextBox = lineForTextBox + "MerchantID --> " + selectedMerchant.mid;
                //if (selectedMerchant.merchantSiteUri != "") lineForTextBox = lineForTextBox + "\nURL -->  " + selectedMerchant.merchantSiteUri;
                //if (selectedMerchant.coupons != "") lineForTextBox = lineForTextBox + "\nCoupons --> " + selectedMerchant.coupons;


                //richTextBox1.Text = lineForTextBox;

            //}

            if (selectedMerchant.merchantSiteUri != null && !Convert.ToString(selectedMerchant.merchantSiteUri).Contains("http")) goToSiteBtn.Enabled = false; else goToSiteBtn.Enabled = true;
            if (selectedMerchant.adminUri != null && !Convert.ToString(selectedMerchant.adminUri).Contains("http")) goToAdminBtn.Enabled = false; else goToAdminBtn.Enabled = true;
            if (selectedMerchant.returnPortalUri != null && !Convert.ToString(selectedMerchant.returnPortalUri).Contains("http")) returnPortalBtn.Enabled = false; else returnPortalBtn.Enabled = true;
            if (selectedMerchant.trackingPortalUri != null && !Convert.ToString(selectedMerchant.trackingPortalUri).Contains("http")) trackingPortalBtn.Enabled = false; else trackingPortalBtn.Enabled = true;


        }

        private void merchantsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            showMerchantDetails(merchantsListBox.SelectedItem.ToString());
        }

        private void goToSiteBtn_Click(object sender, EventArgs e)
        {
            launchUriInChrome(selectedMerchant.merchantSiteUri, selectedMerchant.siteLoginUserName, selectedMerchant.siteLoginPassword);
        }

        private void goToGEAdminBtn_Click(object sender, EventArgs e)
        {
            if (chosenEnvironment == EnvironmentType.QA) { launchUriInChrome(GEAdminQA, "", ""); }
            if (chosenEnvironment == EnvironmentType.Staging) { launchUriInChrome(GEAdminStg, "", ""); }
            if (chosenEnvironment == EnvironmentType.Production) { launchUriInChrome(GEAdminProd, "", ""); }
            
        }

        private void goToAdminBtn_Click(object sender, EventArgs e)
        {
            launchUriInChrome(selectedMerchant.adminUri, selectedMerchant.adminLoginUserName, selectedMerchant.adminLoginPassword);
        }

        private void returnPortalBtn_Click(object sender, EventArgs e)
        {
            launchUriInChrome(selectedMerchant.returnPortalUri, "", "");
        }

        private void trackingPortalBtn_Click(object sender, EventArgs e)
        {
            launchUriInChrome(selectedMerchant.trackingPortalUri, "", "");
        }

        private void launchUriInChrome(String uri, String loginUserName, String loginPassword)
        {

            var options = new ChromeOptions();
            options.AddArgument("incognito");

            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            IWebDriver driver = new ChromeDriver(driverService, options);
            driver.Url = Convert.ToString(uri);

            if (loginUserName != "" && loginPassword != "")
            {
                AutoItX.WinWait("- Google Chrome", "", 1);
                AutoItX.WinActivate("- Google Chrome");
                AutoItX.Send(loginUserName);
                AutoItX.Send("{TAB}", 0);
                AutoItX.Send(loginPassword);
                AutoItX.Send("{TAB}", 0);
                AutoItX.Send("{Enter}", 0);
            }

            driver.Manage().Window.Maximize();

            //try { driver.Navigate().Refresh(); }
            //catch (OpenQA.Selenium.NoSuchWindowException e) { }
        }

        private void QaBtn_Click(object sender, EventArgs e)
        {
            chosenEnvironment = EnvironmentType.QA;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void stagingBtn_Click(object sender, EventArgs e)
        {
            chosenEnvironment = EnvironmentType.Staging;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void productionBtn_Click(object sender, EventArgs e)
        {
            chosenEnvironment = EnvironmentType.Production;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void changeBtnsCollor()
        {
            switch(chosenEnvironment)
            {
                case EnvironmentType.QA:
                    QaBtn.BackColor = Color.LightGreen;
                    stagingBtn.BackColor = Color.Transparent;
                    productionBtn.BackColor = Color.Transparent;
                    break;
                case EnvironmentType.Staging:
                    QaBtn.BackColor = Color.Transparent;
                    stagingBtn.BackColor = Color.LightGreen;
                    productionBtn.BackColor = Color.Transparent;
                    break;
                case EnvironmentType.Production:
                    QaBtn.BackColor = Color.Transparent;
                    stagingBtn.BackColor = Color.Transparent;
                    productionBtn.BackColor = Color.LightGreen;
                    break;
            }
        }


    }
}
