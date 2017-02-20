﻿using Office365Search.Core.Extensions;
using Office365Search.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;


namespace Office365Search.Core.Helpers
{


    public static class SharePointHelper
    {
        public static XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";

        public static async Task<List<string>> GetListNames()
        {
            // List<string> documents = new List<string>();

            var accessToken = await GetAccessTokenForResource("https://kamat777.sharepoint.com/");

            List<string> listnames = new List<string>();

            //string queryText = "nayak";

            string SharePointServiceRoot = "https://kamat777.sharepoint.com/Polaris";
            //StringBuilder requestUri = new StringBuilder().Append(SharePointServiceRoot).Append("/_api/search/query?querytext='" + queryText + "'");
            StringBuilder requestUri = new StringBuilder().Append(SharePointServiceRoot).Append("/_api/web/lists");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUri.ToString());
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            string responseString = await response.Content.ReadAsStringAsync();

            XElement root = XElement.Parse(responseString);

            foreach (XElement entryItem in root.Descendants(d + "Title"))
            {
                listnames.Add(entryItem.Value);
            }

            // MyList.ItemsSource = listnames;



            return listnames;
        }

        public static async Task<List<DocumentInformation>> GetSharePointDocuments(string rootSiteUrl, string searchAPIUrl)
        {
            List<DocumentInformation> documents = new List<DocumentInformation>();

            var accessToken = await GetAccessTokenForResource(rootSiteUrl);

            StringBuilder requestUri = new StringBuilder().Append(rootSiteUrl).Append(searchAPIUrl);

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUri.ToString());
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            string responseString = await response.Content.ReadAsStringAsync();

            XElement root = XElement.Parse(responseString);

            List<XElement> items = root.Descendants(d + "PrimaryQueryResult")
                                     .Elements(d + "RelevantResults")
                                     .Elements(d + "Table")
                                     .Elements(d + "Rows")
                                     .Elements(d + "element")
                                     .ToList();

            foreach (var item in items)
            {
                DocumentInformation document = new DocumentInformation();
                document.Title = item.Element(d + "Cells").Descendants(d + "Key").First(a => a.Value == "Title").Parent.Element(d + "Value").Value;
                // document.AuthorInformation.DisplayName = item.Element(d + "Cells").Descendants(d + "Key").First(a => a.Value == "Author").Parent.Element(d + "Value").Value;
                document.Url = item.Element(d + "Cells").Descendants(d + "Key").First(a => a.Value == "Path").Parent.Element(d + "Value").Value;

                DateTime modifiedDate = Convert.ToDateTime(item.Element(d + "Cells").Descendants(d + "Key").First(a => a.Value == "LastModifiedTime").Parent.Element(d + "Value").Value);
                document.ModifiedDate= modifiedDate;

                documents.Add(document);
            }


            return documents;
        }



        private static async Task<string> GetAccessTokenForResource(string resource)
        {
            string token = null;


            WebAccountProvider aadAccount = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.windows.net");
            // WebTokenRequest request = new WebTokenRequest(aadAccount, string.Empty, App.Current.Resources["ida:ClientId"].ToString(), WebTokenRequestPromptType.Default);
            WebTokenRequest request = new WebTokenRequest(aadAccount, string.Empty, "5535c865-e37e-4dc0-8cc4-726a14feb437", WebTokenRequestPromptType.Default);
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
                request = new WebTokenRequest(aadAccount, string.Empty, "5535c865-e37e-4dc0-8cc4-726a14feb437", WebTokenRequestPromptType.ForceAuthentication);
                request.Properties.Add("authority", "https://login.windows.net");
                request.Properties.Add("resource", resource);

                response = await WebAuthenticationCoreManager.RequestTokenAsync(request);
                if (response.ResponseStatus == WebTokenRequestStatus.Success)
                {
                    WebTokenResponse webToken = response.ResponseData[0];
                    token = webToken.Token;
                }

            }


            return token;
        }


        public static string EnsureEndsWith(this string value, string endsWith)
        {
            return (value.EndsWith(endsWith)) ? value : value + endsWith;
        }

    }
}
