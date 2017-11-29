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
                if (String.IsNullOrEmpty(merchant.mid))
                {
                    String queryMid = buildQueryForReturnSpecificColumnValueFromMerchantsTable("MerchantId", merchant);
                    merchant.mid = DAL.readFromSQL(queryMid, "MerchantId", ConnectionString);
                }

                String queryBrowsingPlatformTmp;
                
                if (!String.IsNullOrEmpty(merchant.mid))
                {
                    queryBrowsingPlatformTmp = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (" +
                    "select top 1 BrowsingPlatformTypeId from Merchants where merchantid = " + merchant.mid + ")";
                }
                else
                {
                    queryBrowsingPlatformTmp = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = ("
                        + buildQueryForReturnSpecificColumnValueFromMerchantsTable("BrowsingPlatformTypeId", merchant);
                }

                String queryBrowsingPlatform = queryBrowsingPlatformTmp;

                //When query has specific characters - replace them to avoid exception
                //
                if (queryBrowsingPlatformTmp.Contains("Paul's"))
                {
                    queryBrowsingPlatform = queryBrowsingPlatformTmp.Replace("Paul's", "Paul_s");
                }

                merchant.browsingPlatformTypeId = DAL.readFromSQL(queryBrowsingPlatform, "MerchantPlatformName", ConnectionString);


                String queryAPIPlatformTmp;

                if (!String.IsNullOrEmpty(merchant.mid))
                {
                    queryAPIPlatformTmp = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = (" +
                    "select top 1 APIPlatformTypeId from Merchants where merchantid = " + merchant.mid + ")";
                }
                else
                {
                    queryAPIPlatformTmp = "select MerchantPlatformName from MerchantPlatforms where MerchantPlatformId = ("
                        + buildQueryForReturnSpecificColumnValueFromMerchantsTable("APIPlatformTypeId", merchant);
                }

                String queryAPIPlatform = queryAPIPlatformTmp;

                //When query has specific characters - replace them to avoid exception
                //
                if (queryBrowsingPlatformTmp.Contains("Paul's"))
                {
                    queryAPIPlatform = queryBrowsingPlatformTmp.Replace("Paul's", "Paul_s");
                }

                merchant.apiPlatformTypeId = DAL.readFromSQL(queryAPIPlatform, "MerchantPlatformName", ConnectionString);

            }

            return merchant;
        }

        private String buildQueryForReturnSpecificColumnValueFromMerchantsTable (String columnName, Merchant merchant)
        {

            return "select top 1 " + columnName + " from Merchants where merchantname like '%" + merchant.merchantName + "%'"
                        + " and IsActive = 1" + " and (SiteURL = '" + merchant.merchantSiteUri + "'"
                        + " or SiteURL = (select LEFT('" + merchant.merchantSiteUri + "', LEN('" + merchant.merchantSiteUri
                        + "')-1)) or SiteURL = (select('" + merchant.merchantSiteUri + "' + '/'))"
                        + " or SiteURL = (select replace('" + merchant.merchantSiteUri
                        + "', 'http://', '')) or SiteURL = (select replace('' + (select LEFT('" + merchant.merchantSiteUri
                        + "', LEN('" + merchant.merchantSiteUri + "')-1)), 'http://', '')) or SiteURL = (select replace('"
                        + merchant.merchantSiteUri + "', 'https://', '')) or SiteURL = (select replace('' + (select LEFT('"
                        + merchant.merchantSiteUri + "', LEN('" + merchant.merchantSiteUri + "')-1)), 'https://', '')))";
        }
    }
}
