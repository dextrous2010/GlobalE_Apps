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

        const String GEAdminUriQA = "https://qa.bglobale.com/GlobaleAdmin";
        const String GEAdminUriStg = "https://www2.bglobale.com/GlobaleAdmin";
        const String GEAdminUriProd = "https://web.global-e.com/GlobaleAdmin";

        EnvironmentType chosenEnvironment = EnvironmentType.QA;

        Merchant selectedMerchant = new Merchant();

        Dictionary<EnvironmentType, EnvironmentData> environmentList = new Dictionary<EnvironmentType, EnvironmentData>();

        static string fileName = string.Empty;
        static string path = string.Empty;


        //Create COM Objects. Create a COM object for everything that is referenced
        static Microsoft.Office.Interop.Excel.Application xlAppQA = new Microsoft.Office.Interop.Excel.Application();
        static Workbook xlWorkbook;


        public GE_Merchant_Picker_Form()
        {

            Console.WriteLine("Trying to upload latest file with merchant addresses from sharepoint...");
            using (var client = new System.Net.WebClient())
            {
                try
                {
                    //client.Credentials = new NetworkCredential("denis.hural@global-e.com", "L0g1tech_10");
                    client.DownloadFile("https://globaleonline-my.sharepoint.com/personal/ifat_perlmandomy_global-e_com/_layouts/15/Doc.aspx?sourcedoc=%7B0fc9c202-737e-438f-975c-8ef3ae822b8d%7D&action=default", "Merchants Adresses.xlsx");
                    fileName = "Merchants Adresses.xlsx";
                    Console.WriteLine("The file has been successfully downloaded.");
                }
                catch
                {
                    Console.WriteLine("Could not download the file. Will work with offline version.");
                    fileName = "Merchants Adresses Offline.xlsx";
                }
            }

            path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" + fileName;

            xlWorkbook = xlAppQA.Workbooks.Open(path);

            InitializeComponent();
            initializeMerchantsListBox();
        }

        void initializeMerchantsListBox()
        {
            switch (chosenEnvironment)
            {
                case EnvironmentType.QA:
                    if (!environmentList.ContainsKey(chosenEnvironment))
                    {
                        System.Action action = () => environmentList.Add(EnvironmentType.QA, new EnvironmentData(5, xlWorkbook.Sheets["QA"], "54.72.115.215"));
                        using (Loading_Form lf = new Loading_Form(action))
                        {
                            lf.ShowDialog(this);
                        }
                    }
                    break;
                case EnvironmentType.Staging:
                    if (!environmentList.ContainsKey(chosenEnvironment))
                    {
                        System.Action action = () => environmentList.Add(EnvironmentType.Staging, new EnvironmentData(5, xlWorkbook.Sheets["Staging"], "54.72.120.2"));
                        using (Loading_Form lf = new Loading_Form(action))
                        {
                            lf.ShowDialog(this);
                        }
                    }
                    break;
                case EnvironmentType.Production:
                    if (!environmentList.ContainsKey(chosenEnvironment))
                    {
                        System.Action action = () => environmentList.Add(EnvironmentType.Production, new EnvironmentData(3, xlWorkbook.Sheets["Production"]));
                        using (Loading_Form lf = new Loading_Form(action))
                        {
                            lf.ShowDialog(this);
                        }
                    }
                    break;
                default:
                    break;
            }

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
            if (!string.IsNullOrWhiteSpace(selectedMerchant.apiPlatformTypeId)) stringBuilder.Append("\nPlatform --> " + selectedMerchant.apiPlatformTypeId);
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
