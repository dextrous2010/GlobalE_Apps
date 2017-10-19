using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GE_Merchant_Picker
{
    public class Merchant
    {
        public String merchantName;
        public String mid;
        public String platformType;
        public String siteLoginUserName;
        public String siteLoginPassword;
        public String adminLoginUserName;
        public String adminLoginPassword;
        public String merchantSiteUri;
        public String adminUri;
        public String returnPortalUri;
        public String trackingPortalUri;
        public String logsUri;
        public String coupons;
        public String comments;

        public Merchant() { }

        public Merchant(String merchantName, String merchantSiteUri)
        {
            this.merchantName = merchantName;
            this.merchantSiteUri = merchantSiteUri;
        }

        public void ResetMerchant()
        {
            merchantName = "";
            mid = "";
            platformType = "";
            siteLoginUserName = "";
            siteLoginPassword = "";
            adminLoginUserName = "";
            adminLoginPassword = "";
            merchantSiteUri = "";
            adminUri = "";
            returnPortalUri = "";
            trackingPortalUri = "";
            logsUri = "";
            coupons = "";
            comments = "";

        }

    }
}
