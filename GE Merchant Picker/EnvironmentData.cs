using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

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
            xlRange = xlWorksheet.UsedRange;

            int tempRow = startRowInExcelWithMerchantName;


            while (xlRange.Cells[tempRow, 1] != null && xlRange.Cells[tempRow, 1].Value2 != null)
            {
                merchantsList.Add(xlRange.Cells[tempRow, 1].Value2.ToString());
                tempRow++;
            }

            if (!string.IsNullOrWhiteSpace(serverIP))
            {
                ConnectionString = "user id=AppUser;" +
                    "password=AppUser;" +
                    "server=" +
                    serverIP +
                    ";" +
                    "Trusted_Connection=no;" +
                    "database=GlobalE;" +
                    "connection timeout=5";
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
                                        + " and SiteURL = '" + merchant.merchantSiteUri + "'" + " and IsActive = 1";
                    merchant.mid = DAL.readFromSQL(queryMid, "MerchantId", ConnectionString);
                }

                String queryPlatformTmp = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (select top 1 MerchantPlatformId from Merchants where merchantname like '%" + merchant.merchantName + "%'"
                                        + " and SiteURL = '" + merchant.merchantSiteUri + "'" + " and IsActive = 1)";

                String queryPlatform = queryPlatformTmp;
                
                //When query has specific characters - replace them to avoid exception
                //
                if (queryPlatformTmp.Contains("Paul's"))
                {
                    queryPlatform = queryPlatformTmp.Replace("Paul's", "Paul_s");
                }

                merchant.apiPlatformTypeId = DAL.readFromSQL(queryPlatform, "MerchantPlatformName", ConnectionString);
                
            }

            return merchant;
        }

    }
}
