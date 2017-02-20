using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Search.Core.Helpers
{
    public static class SettingsHelper
    {
        public static bool SaveSharePointCredentials(string rootSiteUrl, string userName, string password)
        {
            bool successful = false;

            try
            {

                Windows.Storage.ApplicationDataCompositeValue composite = new Windows.Storage.ApplicationDataCompositeValue();

                composite["RootSiteUrl"] = rootSiteUrl;
                composite["UserName"] = userName;
                composite["Password"] = password;

                Windows.Storage.ApplicationDataContainer localSettings = Windows.Storage.ApplicationData.Current.LocalSettings;
                localSettings.Values["SharePointCredentials"] = composite;
            }
            catch (Exception ex)
            {

            }

            return successful;
        }

        public static SharePointCredentialInformation GetSharePointCredentials()
        {
            SharePointCredentialInformation credentialInformation = null;

            Windows.Storage.ApplicationDataCompositeValue composite = null;

            try
            {
                Windows.Storage.ApplicationDataContainer localSettings = Windows.Storage.ApplicationData.Current.LocalSettings;
                composite = (Windows.Storage.ApplicationDataCompositeValue)localSettings.Values["SharePointCredentials"];

                if (composite != null)
                {
                    credentialInformation = new SharePointCredentialInformation()
                    {
                        RootSiteUrl = (string)composite["RootSiteUrl"],
                        UserName = (string)composite["UserName"],
                        Password = (string)composite["Password"],
                    };
                }

            }
            catch (Exception ex)
            {

            }

            return credentialInformation;
        }

        public class SharePointCredentialInformation
        {
            public string RootSiteUrl { get; set; }
            public string UserName { get; set; }
            public string Password { get; set; }
        }


    }

}
