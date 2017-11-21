using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Microsoft.Office.Interop.Excel;
using System.IO;
using AutoIt;
using System.Drawing;
using System.Text;
using System.Net;
using System.Security;
using Microsoft.SharePoint.Client;

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

        static string fileName = "Merchants Adresses.xlsx";
        static string path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\" + fileName);
        string pathToMerchantsFile = Path.Combine(Environment.CurrentDirectory, fileName);


        //Create COM Objects. Create a COM object for everything that is referenced
        static Microsoft.Office.Interop.Excel.Application xlAppQA = new Microsoft.Office.Interop.Excel.Application();
        static Workbook xlWorkbook;

        const string username = "Denis.Hural@global-e.com";
        const string password = "L0g1tech_10";
        const string url = @"https://globaleonline-my.sharepoint.com/:x:/r/personal/ifat_perlmandomy_global-e_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7B0FC9C202-737E-438F-975C-8EF3AE822B8D%7D&file=Merchants%20Adresses.xlsx&action=default&IsList=1&ListId=%7BE4F44CE8-04E4-4662-8FE7-8A1BB9F2F8F0%7D&ListItemId=9";
        

        public GE_Merchant_Picker_Form()
        {
            /*
            using (var client = new System.Net.WebClient())
            {
                client.Credentials = new NetworkCredential("Denis.Hural@global-e.com", "L0g1tech_10");

                //client.UseDefaultCredentials = true;
                //client.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
                String Url = "https://globaleonline-my.sharepoint.com/personal/ifat_perlmandomy_global-e_com/_layouts/15/WopiFrame.aspx?sourcedoc=%7B0FC9C202-737E-438F-975C-8EF3AE822B8D%7D&file=Merchants%20Adresses.xlsx&action=default&IsList=1&ListId=%7BE4F44CE8-04E4-4662-8FE7-8A1BB9F2F8F0%7D&ListItemId=9";
                String destFileName = "Merchants Adresses.xlsx";
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Url);
                req.UserAgent = "testacc";
                client.DownloadFile(Url, destFileName);
            }
            */


            var securedPassword = new SecureString();
            foreach (var c in password.ToCharArray()) securedPassword.AppendChar(c);
            var credentials = new SharePointOnlineCredentials(username, securedPassword);

            DownloadFile(url, credentials, "Merchants Adresses.xlsx");


            xlWorkbook = xlAppQA.Workbooks.Open(path);

            environmentList.Add(EnvironmentType.QA, new EnvironmentData(4, xlWorkbook.Sheets["QA"], "54.72.115.215"));
            environmentList.Add(EnvironmentType.Staging, new EnvironmentData(5, xlWorkbook.Sheets["Staging"], "54.72.120.2"));
            environmentList.Add(EnvironmentType.Production, new EnvironmentData(3, xlWorkbook.Sheets["Production"]));

            InitializeComponent();
            initializeMerchantsListBox();

        }

        private static void DownloadFile(string webUrl, ICredentials credentials, string filePath)
        {
            using (WebClient client = new WebClient())
            {
                //client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                //client.Headers.Add("User-Agent: Other");

                client.UseDefaultCredentials = true;
                //client.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //client.Proxy = null;
                //client.Credentials = credentials;
                //client.Credentials = new NetworkCredential("Denis.Hural@global-e.com", "L0g1tech_10");
                //HttpWebRequest req = (HttpWebRequest)WebRequest.Create(webUrl);
                //req.UserAgent = "testacc";
                client.DownloadFile(new Uri(webUrl), filePath);
            }
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
            if (!string.IsNullOrWhiteSpace(selectedMerchant.platformType)) stringBuilder.Append("\nPlatform --> " + selectedMerchant.platformType);
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
        }
    }
}
