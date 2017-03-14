using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using static Office365Search.Core.Helpers.SettingsHelper;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Office365Search
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            
            this.InitializeComponent();
        }


        private async void GetAccessToken(object sender, RoutedEventArgs e)
        {
            string token = null;
            string resource = "https://kamat777.sharepoint.com/";

            WebAccountProvider aadAccount = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.windows.net");
            // WebTokenRequest request = new WebTokenRequest(aadAccount, string.Empty, App.Current.Resources["ida:ClientId"].ToString(), WebTokenRequestPromptType.Default);
            WebTokenRequest request = new WebTokenRequest(aadAccount, string.Empty, "48d16164-a3b3-4ab1-b626-a5271858fdac", WebTokenRequestPromptType.Default);
            request.Properties.Add("authority", "https://login.windows.net");
            request.Properties.Add("resource", resource);

            var response = await WebAuthenticationCoreManager.GetTokenSilentlyAsync(request);
            if (response.ResponseStatus == WebTokenRequestStatus.Success)
            {
                WebTokenResponse webToken = response.ResponseData[0];
                token = webToken.Token;
            }
            else if (response.ResponseStatus == WebTokenRequestStatus.UserInteractionRequired)
            {
                aadAccount = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.windows.net");
                request = new WebTokenRequest(aadAccount, string.Empty, "48d16164-a3b3-4ab1-b626-a5271858fdac", WebTokenRequestPromptType.ForceAuthentication);
                request.Properties.Add("authority", "https://login.windows.net");
                request.Properties.Add("resource", resource);

                response = await WebAuthenticationCoreManager.RequestTokenAsync(request);

                if (response.ResponseStatus == WebTokenRequestStatus.Success)
                {
                    WebTokenResponse webToken = response.ResponseData[0];
                    token = webToken.Token;
                }

            }

            string tokenfound = token;






        }


        private void OnAuthorizeSharePointClick(object sender, RoutedEventArgs e)
        {
            StartSharePointAuthorization(sender as Button);
        }

        private async void StartSharePointAuthorization(Button button)
        {
            bool isAuthorized = false;
            button.IsEnabled = false;

            string siteUrl = sharePointSiteUrlTextBox.Text.Trim();
            string userName = sharePointUserNameTextBox.Text.Trim();
            string password = sharePointPasswordTextBox.Password.Trim();

            SharePointCredentialInformation credentials = Core.Helpers.SettingsHelper.GetSharePointCredentials();

            if (credentials == null)
            {
                isAuthorized = await Core.Helpers.ContextHelper.ValidateCredentialsAsync(siteUrl, userName, password);

                if (isAuthorized)
                {
                    Core.Helpers.SettingsHelper.SaveSharePointCredentials(siteUrl, userName, password);
                }

            }

            button.IsEnabled = credentials == null;
        }


    }
}
