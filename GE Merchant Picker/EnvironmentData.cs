using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GE_Merchant_Picker
{
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

            if (!string.IsNullOrWhiteSpace(ConnectionString))
            {
                if (merchant.mid == "")
                {

                    String queryMid = "select top 1 MerchantId from Merchants where merchantname like '%" + merchant.merchantName + "%'"
                                        + " and IsActive = 1 and SiteURL = '" + merchant.merchantSiteUri + "'";
                    merchant.mid = DAL.readFromSQL(queryMid, "MerchantId", ConnectionString);
                }

                String queryPlatform = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (select top 1 MerchantPlatformId from Merchants where merchantname like '%"
                                    + merchant.merchantName + "%' and IsActive = 1)";
                merchant.platformType = DAL.readFromSQL(queryPlatform, "MerchantPlatformName", ConnectionString);
            }


            String lineForTextBox = "";

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
}
