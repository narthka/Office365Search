using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Search.Core.Helpers
{
   public class ContextHelper
    {

        public async static Task<bool> ValidateCredentialsAsync(string siteUrl, string userName, string password)
        {
            bool successful = false;

            try
            {
                using (ClientContext context = GetCredentialedContext(siteUrl, userName, password))
                {
                    Web web = context.Web;

                    context.Load(web, w => w.Title);

                    await context.ExecuteQueryAsync();

                    successful = web.Title != null;

                }
            }
            catch (Exception ex)
            {

            }

            return successful;
        }



        private static ClientContext GetCredentialedContext(string siteUrl, string userName, string password)
        {
            return new ClientContext(siteUrl)
            {
                Credentials = new SharePointOnlineCredentials(userName, password),
            };
        }


    }
}
