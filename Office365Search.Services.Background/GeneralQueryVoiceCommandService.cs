using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Office365Search.Core.Extensions;
using Office365Search.Core.Helpers;
using Office365Search.Core.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Background;
using Windows.ApplicationModel.Resources.Core;
using Windows.ApplicationModel.VoiceCommands;
using Windows.Media.SpeechRecognition;

namespace Office365Search.Services.Background
{
    public sealed class GeneralQueryVoiceCommandService : IBackgroundTask
    {

        VoiceCommandServiceConnection voiceServiceConnection;
        BackgroundTaskDeferral serviceDeferral;
        ResourceMap cortanaResourceMap;
        ResourceContext cortanaContext;
        DateTimeFormatInfo dateFormatInfo;

        public async void Run(IBackgroundTaskInstance taskInstance)
        {
            serviceDeferral = taskInstance.GetDeferral();

            // Register to receive an event if Cortana dismisses the background task. This will
            // occur if the task takes too long to respond, or if Cortana's UI is dismissed.
            // Any pending operations should be cancelled or waited on to clean up where possible.
            taskInstance.Canceled += OnTaskCanceled;

            var triggerDetails = taskInstance.TriggerDetails as AppServiceTriggerDetails;

            // Load localized resources for strings sent to Cortana to be displayed to the user.
            cortanaResourceMap = ResourceManager.Current.MainResourceMap.GetSubtree("Resources");

            // Select the system language, which is what Cortana should be running as.
            cortanaContext = ResourceContext.GetForViewIndependentUse();

            // Get the currently used system date format
            dateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

            // This should match the uap:AppService and VoiceCommandService references from the 
            // package manifest and VCD files, respectively. Make sure we've been launched by
            // a Cortana Voice Command.
            if (triggerDetails != null && triggerDetails.Name == "GeneralQueryVoiceCommandService")
            {
                try
                {
                    voiceServiceConnection =
                        VoiceCommandServiceConnection.FromAppServiceTriggerDetails(
                            triggerDetails);

                    voiceServiceConnection.VoiceCommandCompleted += OnVoiceCommandCompleted;

                    // GetVoiceCommandAsync establishes initial connection to Cortana, and must be called prior to any 
                    // messages sent to Cortana. Attempting to use ReportSuccessAsync, ReportProgressAsync, etc
                    // prior to calling this will produce undefined behavior.
                    VoiceCommand voiceCommand = await voiceServiceConnection.GetVoiceCommandAsync();
                    var interpretation = voiceCommand.SpeechRecognitionResult.SemanticInterpretation;

                    string clientId = cortanaResourceMap.GetValue("ClientId", cortanaContext).ValueAsString;
                    string userName = cortanaResourceMap.GetValue("Domain", cortanaContext).ValueAsString;
                    string rootSiteUrl = cortanaResourceMap.GetValue("rootSite", cortanaContext).ValueAsString;

                    StringBuilder searchAPIUrl = new StringBuilder();

                    switch (voiceCommand.CommandName)
                    {
                        case "SharePointWhatsCheckedOutQueryCommand":
                            searchAPIUrl = searchAPIUrl.Append("/_api/search/query?querytext='CheckoutUserOWSUSER:" + userName + "'");
                            await SearchCheckedOutDocumentsAsync(rootSiteUrl, searchAPIUrl.ToString());
                            break;

                        case "SPSearchContentCommand":

                            var searchSiteName = voiceCommand.Properties["searchsite"][0];
                            var searchText = voiceCommand.Properties["dictatedSearchText"][0];
                            searchAPIUrl = searchAPIUrl.Append("/_api/search/query?querytext='" + searchText + "'");
                            await SearchSharePointDocumentsAsync(rootSiteUrl, searchAPIUrl.ToString());
                            break;

                        default:
                            // As with app activation VCDs, we need to handle the possibility that
                            // an app update may remove a voice command that is still registered.
                            // This can happen if the user hasn't run an app since an update.
                            LaunchAppInForeground();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Handling Voice Command failed " + ex.ToString());
                }
            }
        }


        private async Task SearchCheckedOutDocumentsAsync(string rootSiteUrl, string searchAPIUrl)
        {
            await ShowProgressScreen("Finding documents checked out to you...");

            VoiceCommandResponse response;
            var destinationsContentTiles = new List<VoiceCommandContentTile>();

            var documents = await Core.Helpers.SharePointHelper.GetSharePointDocuments(rootSiteUrl, searchAPIUrl);

            if (documents.Count > 0)
            {
                foreach (var document in documents)
                {
                    var destinationTile = new VoiceCommandContentTile();

                    try
                    {
                        //destinationTile.ContentTileType = VoiceCommandContentTileType.TitleWith68x68IconAndText;
                        // destinationTile.Image = await Windows.Storage.StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///OfficePoint.Services.Background" + document.IconUrl.EnsureStartsWith("/")));
                        // destinationTile.AppContext = document;
                        // destinationTile.AppLaunchArgument = "type=" + "SharePointWhatsCheckedOutQueryCommand" + "&itemId=" + document.ItemId.ToString();

                        destinationTile.Title = document.Title;
                        destinationTile.TextLine1 = "Last modified: " + document.ModifiedDate.ToString();
                        //  destinationTile.TextLine1 = document.AuthorInformation.DisplayName;
                        // destinationTile.TextLine1 = "modified: " + document;
                        // destinationTile.TextLine2 = "views: " + document.ViewCount;


                        destinationsContentTiles.Add(destinationTile);
                    }
                    catch (Exception ex)
                    {

                    }
                }

                await ShowProgressScreen("I found " + documents.Count + " documents...");
                response = VoiceCommandResponse.CreateResponse(new VoiceCommandUserMessage()
                {
                    DisplayMessage = "Here are your checked out documents",
                    SpokenMessage = "Here are your top " + destinationsContentTiles.Count + " checked out documents"

                }, destinationsContentTiles);

                await voiceServiceConnection.ReportSuccessAsync(response);
            }
            else
            {
                response = VoiceCommandResponse.CreateResponse(new VoiceCommandUserMessage()
                {
                    DisplayMessage = "There's nothing checked out to you",
                    SpokenMessage = "I didn't find anything checked out to you. Time to get working on something."

                }, destinationsContentTiles);

                await voiceServiceConnection.ReportSuccessAsync(response);
            }


            return;
        }


        private async Task SearchSharePointDocumentsAsync(string rootSiteUrl, string searchAPIUrl)
        {
            await ShowProgressScreen("Searching Documents...");

            VoiceCommandResponse response;
            var destinationsContentTiles = new List<VoiceCommandContentTile>();

            var documents = await Core.Helpers.SharePointHelper.GetSharePointDocuments(rootSiteUrl, searchAPIUrl);

            if (documents.Count > 0)
            {
                foreach (var document in documents.OrderByDescending(d=>d.ModifiedDate))
                {
                    var destinationTile = new VoiceCommandContentTile();

                    try
                    {
                        destinationTile.ContentTileType = VoiceCommandContentTileType.TitleWithText;
                        // destinationTile.Image = await Windows.Storage.StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///OfficePoint.Services.Background" + document.IconUrl.EnsureStartsWith("/")));
                        // destinationTile.AppContext = document;
                        // destinationTile.AppLaunchArgument = "type=" + "SharePointWhatsCheckedOutQueryCommand" + "&itemId=" + document.ItemId.ToString();
                        destinationTile.Title = document.Title;
                        destinationTile.TextLine1 = "Last modified: " + document.ModifiedDate.ToString();
                        //  destinationTile.TextLine1 = document.AuthorInformation.DisplayName;
                        // destinationTile.TextLine1 = "modified: " + document;
                        // destinationTile.TextLine2 = "views: " + document.ViewCount;

                        destinationsContentTiles.Add(destinationTile);
                    }
                    catch (Exception ex)
                    {

                    }
                }

                await ShowProgressScreen("I found " + documents.Count + " documents...");
                response = VoiceCommandResponse.CreateResponse(new VoiceCommandUserMessage()
                {
                    DisplayMessage = "Here are the documents",
                    SpokenMessage = "Here are your " + destinationsContentTiles.Count + " documents"

                }, destinationsContentTiles);

                await voiceServiceConnection.ReportSuccessAsync(response);
            }
            else
            {
                response = VoiceCommandResponse.CreateResponse(new VoiceCommandUserMessage()
                {
                    DisplayMessage = "No documents found for the given Search term",
                    SpokenMessage = "I didn't find anything with given search term."

                }, destinationsContentTiles);

                await voiceServiceConnection.ReportSuccessAsync(response);
            }


            return;
        }



        private async Task SharePointSearchContentInBrowser(string searchSite, string searchText)
        {

            if (!string.IsNullOrEmpty(searchSite))
            {
                searchSite = searchSite.ToLower();
            }

            if (!string.IsNullOrEmpty(searchText))
            {
                searchText = searchText.ToLower();
            }

            //var sharePointCredentials = Core.Helpers.SettingsHelper.GetSharePointCredentials();

            //if (sharePointCredentials != null)
            //{
            //    string url = sharePointCredentials.RootSiteUrl;
            //}
            string uriToLaunch = @"https://www.google.com";


            switch (searchSite)
            {
                case "google":
                    uriToLaunch = @"https://www.google.com";
                    break;

                case "bing":
                    uriToLaunch = @"https://www.bing.com";
                    break;

                case "polaris":
                    uriToLaunch = @"https://www.bing.com";
                    break;


                case "sharepoint online":
                    uriToLaunch = @"https://kamat777.sharepoint.com/_layouts/15/osssearchresults.aspx";
                    break;


                case "insideemc":
                    uriToLaunch = @"https://www.bing.com";
                    break;


                default:
                    break;
            }

            uriToLaunch = uriToLaunch + "/search?q=test";

            // Create a Uri object from a URI string 
            var uri = new Uri(uriToLaunch);

            // Launch the URI
            var success = await Windows.System.Launcher.LaunchUriAsync(uri);

            var userMessage = new VoiceCommandUserMessage();
            string message = string.Empty;
            if (success)
            {
                message = "Search results are displayed in Browser";
            }
            else
            {
                message = "No Search results displayed";
            }

            userMessage.DisplayMessage = message;
            userMessage.SpokenMessage = message;

            var response = VoiceCommandResponse.CreateResponse(userMessage);
            response.AppLaunchArgument = "";
            await voiceServiceConnection.ReportSuccessAsync(response);




            //List<SearchItem> SharePointList = new List<SearchItem>();
            //var sharePointCredentials = Core.Helpers.SettingsHelper.GetSharePointCredentials();

            //if (sharePointCredentials != null)
            //{
            //    string url = sharePointCredentials.RootSiteUrl;
            //    using (ClientContext clientContext = new ClientContext(url))
            //    {
            //        clientContext.Credentials = new SharePointOnlineCredentials(sharePointCredentials.UserName, sharePointCredentials.Password);
            //        clientContext.AuthenticationMode = ClientAuthenticationMode.Default;

            //        List announcementsList = clientContext.Web.Lists.GetByTitle("Announcements");

            //        CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
            //        ListItemCollection items = announcementsList.GetItems(query);




            //        // Web web = clientContext.Web;
            //        // clientContext.Load(web);
            //        clientContext.Load(items);
            //        var webTask = clientContext.ExecuteQueryAsync();

            //        webTask.Wait();


            //        foreach (ListItem listItem in items)
            //        {
            //            SharePointList.Add(new SearchItem() { Title = listItem["Title"].ToString(), URL = "" });
            //            // SharePointList.Add(listItem["Title"].ToString());
            //        }

            //    }


            //}


            //IEnumerable<SearchItem> searchresults = SharePointList.Where(s => s.Title.ToLower() == searchText.ToLower()).ToList();

            //var userMessage = new VoiceCommandUserMessage();
            //var searchResultContentTiles = new List<VoiceCommandContentTile>();

            //if (searchresults.Count() == 0)
            //{
            //    // In this scenario, perhaps someone has modified data on your service outside of your 
            //    // control. If you're accessing a remote service, having a background task that
            //    // periodically refreshes the phrase list so it's likely to be in sync is ideal.
            //    // This is unlikely to occur for this sample app, however.
            //    string foundNosearchResults = "No Results found for search keyword " + searchText;
            //    userMessage.DisplayMessage = foundNosearchResults;
            //    userMessage.SpokenMessage = foundNosearchResults;
            //}

            //else
            //{
            //    // Set a title message for the page.
            //    string message = "";
            //    if (searchresults.Count() > 1)
            //    {
            //        message = searchresults.Count() + " Results found that match " + searchText;
            //    }
            //    else
            //    {
            //        message = "One Result found that match " + searchText;
            //    }
            //    userMessage.DisplayMessage = message;
            //    userMessage.SpokenMessage = message;

            //    // file in tiles for each destination, to display information about the trips without
            //    // launching the app.
            //    foreach (SearchItem item in searchresults)
            //    {
            //        int i = 1;

            //        var searchResultContentTile = new VoiceCommandContentTile();

            //        // To handle UI scaling, Cortana automatically looks up files with FileName.scale-<n>.ext formats based on the requested filename.
            //        // See the VoiceCommandService\Images folder for an example.
            //        // searchResultContentTile.ContentTileType = VoiceCommandContentTileType.TitleWith68x68IconAndText;
            //        //searchResultContentTile.Image = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///AdventureWorks.VoiceCommands/Images/GreyTile.png"));

            //        searchResultContentTile.AppLaunchArgument = item.URL;
            //        searchResultContentTile.Title = item.Title;
            //        //if (trip.StartDate != null)
            //        //{
            //        //    destinationTile.TextLine1 = trip.StartDate.Value.ToString(dateFormatInfo.LongDatePattern);
            //        //}
            //        //else
            //        //{
            //        //    destinationTile.TextLine1 = trip.Destination + " " + i;
            //        //}

            //        searchResultContentTiles.Add(searchResultContentTile);
            //        i++;
            //    }


            //}


            //var response = VoiceCommandResponse.CreateResponse(userMessage, searchResultContentTiles);

            //if (searchresults.Count() > 0)
            //{
            //    response.AppLaunchArgument = "";
            //}

            //await voiceServiceConnection.ReportSuccessAsync(response);




        }

        private async Task ShowProgressScreen(string message)
        {
            var userProgressMessage = new VoiceCommandUserMessage();
            userProgressMessage.DisplayMessage = userProgressMessage.SpokenMessage = message;

            VoiceCommandResponse response = VoiceCommandResponse.CreateResponse(userProgressMessage);
            await voiceServiceConnection.ReportProgressAsync(response);
        }

        private async void LaunchAppInForeground()
        {
            var userMessage = new VoiceCommandUserMessage();
            userMessage.SpokenMessage = "Launching Adventure Works";

            var response = VoiceCommandResponse.CreateResponse(userMessage);

            response.AppLaunchArgument = "";

            await voiceServiceConnection.RequestAppLaunchAsync(response);
        }

        private void OnVoiceCommandCompleted(VoiceCommandServiceConnection sender, VoiceCommandCompletedEventArgs args)
        {
            if (this.serviceDeferral != null)
            {
                this.serviceDeferral.Complete();
            }
        }

        private void OnTaskCanceled(IBackgroundTaskInstance sender, BackgroundTaskCancellationReason reason)
        {
            System.Diagnostics.Debug.WriteLine("Task cancelled, clean up");
            if (this.serviceDeferral != null)
            {
                //Complete the service deferral
                this.serviceDeferral.Complete();
            }
        }


    }
}
