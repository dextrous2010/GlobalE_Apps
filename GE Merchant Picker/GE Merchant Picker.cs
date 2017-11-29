using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using System.IO;
using AutoIt;
using System.Drawing;
using System.Text;
using System.Net;
using System.Net.NetworkInformation;

namespace GE_Merchant_Picker
{
    enum EnvironmentType
    {
        QA,
        Staging,
        Production
    }

    public partial class GE_Merchant_Picker_Form : System.Windows.Forms.Form
    {

        const String GEAdminUriQA = "https://qa.bglobale.com/GlobaleAdmin";
        const String GEAdminUriStg = "https://www2.bglobale.com/GlobaleAdmin";
        const String GEAdminUriProd = "https://web.global-e.com/GlobaleAdmin";

        EnvironmentType chosenEnvironment = EnvironmentType.QA;

        Merchant selectedMerchant = new Merchant();

        Dictionary<EnvironmentType, EnvironmentData> environmentList = new Dictionary<EnvironmentType, EnvironmentData>();

        const string fileName = "Merchants Adresses.xlsx";
        string pathToMerchantsFile = Path.Combine(Environment.CurrentDirectory, fileName);
        const string urlToDownloadMerhcnatsFile = @"https://globaleonline-my.sharepoint.com/personal/ifat_perlmandomy_global-e_com/_layouts/15/download.aspx?guestaccesstoken=kLW64SzxxAp0WOSrhQYUfiY2mtfj8kuh3auEBOYDy4c%3D&docid=10fc9c202737e438f975c8ef3ae822b8d&rev=1&e=0ab994c81d334fbb80357e0d026af2f6";

        //Create COM Objects. Create a COM object for everything that is referenced
        static Microsoft.Office.Interop.Excel.Application xlAppQA = new Microsoft.Office.Interop.Excel.Application();
        static Workbook xlWorkbook;


        public GE_Merchant_Picker_Form()
        {
            //Get latest Merchant's file from sharepoint
            using (var client = new System.Net.WebClient())
            {
                //String destFileName = "Merchants Adresses.xlsx";
                client.UseDefaultCredentials = true;
                try
                {
                    client.DownloadFile(urlToDownloadMerhcnatsFile, pathToMerchantsFile);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Couldn't download the latest Merchants Adresses file from the server.\nContinuing to work with the local version.");
                }
            }

            xlWorkbook = xlAppQA.Workbooks.Open(pathToMerchantsFile);

            environmentList.Add(EnvironmentType.QA, new EnvironmentData(5, xlWorkbook.Sheets["QA"], "54.72.115.215"));
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
                MessageBox.Show("Can't read data from DB");
            }

            StringBuilder stringBuilder = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(selectedMerchant.mid)) stringBuilder.Append("MerchantID --> " + selectedMerchant.mid);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.browsingPlatformTypeId)) stringBuilder.Append("\nBrowsing Platform --> " + selectedMerchant.browsingPlatformTypeId);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.apiPlatformTypeId)) stringBuilder.Append("\nAPI Platform --> " + selectedMerchant.apiPlatformTypeId);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.merchantSiteUri)) stringBuilder.Append("\nURL -->  " + selectedMerchant.merchantSiteUri);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.siteLoginUserName)) stringBuilder.Append("\nUser -->  " + selectedMerchant.siteLoginUserName);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.siteLoginPassword)) stringBuilder.Append("\nPass -->  " + selectedMerchant.siteLoginPassword);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.adminUri)) stringBuilder.Append("\nAdmin --> " + selectedMerchant.adminUri);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.adminLoginUserName)) stringBuilder.Append("\nUser -->  " + selectedMerchant.adminLoginUserName);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.adminLoginPassword)) stringBuilder.Append("\nPass -->  " + selectedMerchant.adminLoginPassword);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.returnPortalUri)) stringBuilder.Append("\nRetrun Portal --> " + selectedMerchant.returnPortalUri);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.trackingPortalUri)) stringBuilder.Append("\nTracking Portal --> " + selectedMerchant.trackingPortalUri);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.coupons)) stringBuilder.Append("\nCoupons --> " + selectedMerchant.coupons);
            if (!string.IsNullOrWhiteSpace(selectedMerchant.comments)) stringBuilder.Append("\nComment --> " + selectedMerchant.comments);

            richTextBox1.Text = stringBuilder.ToString();

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
            if (chosenEnvironment == EnvironmentType.QA) { launchUriInChrome(GEAdminUriQA, "", ""); }
            if (chosenEnvironment == EnvironmentType.Staging) { launchUriInChrome(GEAdminUriStg, "", ""); }
            if (chosenEnvironment == EnvironmentType.Production) { launchUriInChrome(GEAdminUriProd, "", ""); }

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
            options.AddArguments("start-maximized");

            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            //In case user will imediatelly close the browser the following line will throw an exeption
            //to avoid it just catch this exeption
            try
            {
                IWebDriver driver = new ChromeDriver(driverService, options);
                driver.Url = Convert.ToString(uri);


                if (!String.IsNullOrWhiteSpace(loginUserName) && !String.IsNullOrWhiteSpace(loginPassword))
                {
                    AutoItX.WinWait("- Google Chrome", "", 1);
                    AutoItX.WinActivate("- Google Chrome");

                    //Put credentials into autorization popup
                    AutoItX.Send(loginUserName + "{TAB}" + loginPassword + "{TAB}");

                    //Put credentials into autorization popup with confirmation
                    //AutoItX.Send(loginUserName + "{TAB}" + loginPassword + "{TAB}" + "{Enter}");
                }

            }
            catch (System.InvalidOperationException e) { }

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
            //alert.SetAuthenticationCredentials(loginUserName, loginPassword);

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
            switch (chosenEnvironment)
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

        private void GE_Merchant_Picker_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            xlWorkbook.Close();

            //Kills stray local chromedriver.exe instances.
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "kill_chromedriver.bat";
            process.Start();
        }
    }
}
