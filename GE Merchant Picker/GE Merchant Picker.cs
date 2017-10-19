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

    public partial class GE_Merchant_Picker_Form : Form
    {
        
        const String GEAdminQA = "https://qa.bglobale.com/GlobaleAdmin";
        const String GEAdminStg = "https://www2.bglobale.com/GlobaleAdmin";
        const String GEAdminProd = "https://web.global-e.com/GlobaleAdmin";

        bool firstRun = true;
        EnvironmentType chosenEnvironment = EnvironmentType.QA;

        Merchant selectedMerchant = new Merchant();

        Dictionary<EnvironmentType, EnvironmentData> environmentList = new Dictionary<EnvironmentType, EnvironmentData>();

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
            try
            {
                selectedMerchant = environmentList[chosenEnvironment].GetMerchant(merchant);
            }
            catch (Exception)
            {
                selectedMerchant.ResetMerchant();
                MessageBox.Show("Can't take read data from DB");
            }

                String lineForTextBox = "";
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
