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

namespace GE_Merchant_Picker
{
    public partial class GE_Merchant_Picker_Form : Form
    {
        
        String GEAdminQA = "https://qa.bglobale.com/GlobaleAdmin";
        String GEAdminStg = "https://www2.bglobale.com/GlobaleAdmin";
        String GEAdminProd = "https://web.global-e.com/GlobaleAdmin";

        bool firstRun = true;
        int chosenEnvironment = 0;

        Merchant selectedMerchant = new Merchant();

        //Define the row from which all merchant names is started
        static int startRowInExcelWithMerchantNameQA = 4;
        static int startRowInExcelWithMerchantNameStaging = 5;
        static int startRowInExcelWithMerchantNameProd = 3;

        List<String> merchantsListQA = new List<string>();
        List<String> merchantsListStg = new List<string>();
        List<String> merchantsListProd = new List<string>();

        static string fileName = "Auto Merchants Adresses.xlsx";
        static string path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\" + fileName);

        //Create COM Objects. Create a COM object for everything that is referenced
        static Microsoft.Office.Interop.Excel.Application xlAppQA = new Microsoft.Office.Interop.Excel.Application();
        static Workbook xlWorkbook = xlAppQA.Workbooks.Open(path);

        static Worksheet xlWorksheetQA = xlWorkbook.Sheets["QA"];
        static Range xlRangeQA = xlWorksheetQA.UsedRange;

        static Worksheet xlWorksheetStg = xlWorkbook.Sheets["Staging"];
        static Range xlRangeStg = xlWorksheetStg.UsedRange;

        static Worksheet xlWorksheetProd = xlWorkbook.Sheets["Production"];
        static Range xlRangeProd = xlWorksheetProd.UsedRange;

        public GE_Merchant_Picker_Form()
        {

            InitializeComponent();
            initializeMerchantsListBox();

        }

        public void initializeMerchantsListBox()
        {

            if (firstRun)
            {
                //Initialize QA merchants
                //
                int tempRow = startRowInExcelWithMerchantNameQA;

                while (xlRangeQA.Cells[tempRow, 1] != null && xlRangeQA.Cells[tempRow, 1].Value2 != null)
                {
                    merchantsListQA.Add(xlRangeQA.Cells[tempRow, 1].Value2.ToString());
                    tempRow++;
                }


                //Initialize Staging merchants
                //
                tempRow = startRowInExcelWithMerchantNameStaging;

                while (xlRangeStg.Cells[tempRow, 1] != null && xlRangeStg.Cells[tempRow, 1].Value2 != null)
                {
                    merchantsListStg.Add(xlRangeStg.Cells[tempRow, 1].Value2.ToString());
                    tempRow++;
                }

                //Initialize Production merchants
                //
                tempRow = startRowInExcelWithMerchantNameProd;

                while (xlRangeProd.Cells[tempRow, 1] != null && xlRangeProd.Cells[tempRow, 1].Value2 != null)
                {
                    merchantsListProd.Add(xlRangeProd.Cells[tempRow, 1].Value2.ToString());
                    tempRow++;
                }

                merchantsListBox.DataSource = merchantsListQA;

                firstRun = false;
            }

            if (chosenEnvironment == 0) merchantsListBox.DataSource = merchantsListQA;
            else if (chosenEnvironment == 1) merchantsListBox.DataSource = merchantsListStg;
            else if (chosenEnvironment == 2) merchantsListBox.DataSource = merchantsListProd;
        }

        public void showMerchantDetails(string merchant)
        {

            int tempRow = startRowInExcelWithMerchantNameQA;
            Range xlRange = xlRangeQA;

            selectedMerchant.ResetMerchant();

            if (chosenEnvironment != 2)
            {
                if (chosenEnvironment == 0)
                {
                    tempRow = merchantsListQA.IndexOf(merchant) + startRowInExcelWithMerchantNameQA;
                    xlRange = xlRangeQA;
                }
                else if (chosenEnvironment == 1)
                {
                    tempRow = merchantsListStg.IndexOf(merchant) + startRowInExcelWithMerchantNameStaging;
                    xlRange = xlRangeStg;
                }

                //Initialize selected merchant
                //
                //selectedMerchant.merchantSiteUri = xlRange.Cells[tempRow, 2].Value2 != null ? xlRange.Cells[tempRow, 2].Value2.ToString() : "";

                selectedMerchant.merchantName = (xlRange.Cells[tempRow, 1].Value2 ?? String.Empty).ToString();
                selectedMerchant.merchantSiteUri = (xlRange.Cells[tempRow, 2].Value2 ?? String.Empty).ToString();
                selectedMerchant.adminUri = (xlRange.Cells[tempRow, 3].Value2 ?? String.Empty).ToString();
                selectedMerchant.adminLoginUserName = (xlRange.Cells[tempRow, 4].Value2 ?? String.Empty).ToString();
                selectedMerchant.adminLoginPassword = (xlRange.Cells[tempRow, 5].Value2 ?? String.Empty).ToString();
                selectedMerchant.mid = (xlRange.Cells[tempRow, 6].Value2 ?? String.Empty).ToString();
                selectedMerchant.siteLoginUserName = (xlRange.Cells[tempRow, 7].Value2 ?? String.Empty).ToString();
                selectedMerchant.siteLoginPassword = (xlRange.Cells[tempRow, 8].Value2 ?? String.Empty).ToString();

                selectedMerchant.comments = (xlRange.Cells[tempRow, 9].Value2 ?? String.Empty).ToString();
                selectedMerchant.returnPortalUri = (xlRange.Cells[tempRow, 10].Value2 ?? String.Empty).ToString();
                selectedMerchant.logsUri = (xlRange.Cells[tempRow, 11].Value2 ?? String.Empty).ToString();
                selectedMerchant.coupons = (xlRange.Cells[tempRow, 12].Value2 ?? String.Empty).ToString();
                selectedMerchant.trackingPortalUri = (xlRange.Cells[tempRow, 13].Value2 ?? String.Empty).ToString();

                if (selectedMerchant.mid == "")
                {

                    String queryMid = "select top 1 MerchantId from Merchants where merchantname like '%" + selectedMerchant.merchantName + "%'"
                                        + " and IsActive = 1 and SiteURL = '" + selectedMerchant.merchantSiteUri + "'";
                    selectedMerchant.mid = readFromSQL(queryMid, "MerchantId");
                }

                String queryPlatform = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (select top 1 MerchantPlatformId from Merchants where merchantname like '%"
                                    + selectedMerchant.merchantName + "%' and IsActive = 1)";
                selectedMerchant.platformType = readFromSQL(queryPlatform, "MerchantPlatformName");

                String lineForTextBox = "";

                if (selectedMerchant.mid != "") lineForTextBox = lineForTextBox + "MerchantID --> " + selectedMerchant.mid;
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

            }
            else if (chosenEnvironment == 2)
            {

                tempRow = merchantsListProd.IndexOf(merchant) + startRowInExcelWithMerchantNameProd;
                xlRange = xlRangeProd;

                selectedMerchant.merchantName = (xlRange.Cells[tempRow, 1].Value2 ?? String.Empty).ToString();
                selectedMerchant.merchantSiteUri = (xlRange.Cells[tempRow, 2].Value2 ?? String.Empty).ToString();
                selectedMerchant.mid = (xlRange.Cells[tempRow, 3].Value2 ?? String.Empty).ToString();
                selectedMerchant.coupons = (xlRange.Cells[tempRow, 4].Value2 ?? String.Empty).ToString();

                String lineForTextBox = "";

                if (selectedMerchant.mid != "") lineForTextBox = lineForTextBox + "MerchantID --> " + selectedMerchant.mid;
                if (selectedMerchant.merchantSiteUri != "") lineForTextBox = lineForTextBox + "\nURL -->  " + selectedMerchant.merchantSiteUri;
                if (selectedMerchant.coupons != "") lineForTextBox = lineForTextBox + "\nCoupons --> " + selectedMerchant.coupons;


                richTextBox1.Text = lineForTextBox;

            }

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
            if (chosenEnvironment == 0) { launchUriInChrome(GEAdminQA, "", ""); }
            if (chosenEnvironment == 1) { launchUriInChrome(GEAdminStg, "", ""); }
            if (chosenEnvironment == 2) { launchUriInChrome(GEAdminProd, "", ""); }
            
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
            chosenEnvironment = 0;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void stagingBtn_Click(object sender, EventArgs e)
        {
            chosenEnvironment = 1;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void productionBtn_Click(object sender, EventArgs e)
        {
            chosenEnvironment = 2;
            changeBtnsCollor();
            initializeMerchantsListBox();
        }

        private void changeBtnsCollor()
        {
            switch(chosenEnvironment)
            {
                case 0:
                    QaBtn.BackColor = Color.LightGreen;
                    stagingBtn.BackColor = Color.Transparent;
                    productionBtn.BackColor = Color.Transparent;
                    break;
                case 1:
                    QaBtn.BackColor = Color.Transparent;
                    stagingBtn.BackColor = Color.LightGreen;
                    productionBtn.BackColor = Color.Transparent;
                    break;
                case 2:
                    QaBtn.BackColor = Color.Transparent;
                    stagingBtn.BackColor = Color.Transparent;
                    productionBtn.BackColor = Color.LightGreen;
                    break;
            }
        }

        public String readFromSQL(String query, String columnName)
        {
            String result = "";
            SqlConnection myConnection = new SqlConnection();

            if (chosenEnvironment == 0)
            {
                myConnection = new SqlConnection("user id=sql_qa_ukr_DenisH;" +
                           "password=Admin_141;server=54.72.115.215;" +
                           "Trusted_Connection=no;" +
                           "database=GlobalE; " +
                           "connection timeout=30");
            }

            if (chosenEnvironment == 1)
            {
                myConnection = new SqlConnection("user id=sql_qa_ukr_DenisH;" +
                            "password=Admin_141;server=54.72.120.2;" +
                            "Trusted_Connection=no;" +
                            "database=GlobalE; " +
                            "connection timeout=30");
            }

            try
            {
                myConnection.Open();
            }
            catch (Exception) { }

            try
            {
                SqlDataReader myReader = null;
                SqlCommand myCommand = new SqlCommand(query, myConnection);
                myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    result = myReader[columnName].ToString();
                }
            }
            catch (Exception) { }

            try
            {
                myConnection.Close();
            }
            catch (Exception) { }

            return result;
        }

        private void GE_Merchant_Picker_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            xlWorkbook.Close();
        }


    }
}
