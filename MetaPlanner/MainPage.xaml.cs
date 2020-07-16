using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using Windows.UI.Popups;
using MetaPlanner.Model;
using MetaPlanner.Output;
using System.IO;
using Windows.Storage;
using System.Text;
using Serilog;



// La plantilla de elemento Página en blanco está documentada en https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0xc0a

namespace MetaPlanner
{
    /// <summary>
    /// Página vacía que se puede usar de forma independiente o a la que se puede navegar dentro de un objeto Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        public static AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile();

        //Set the scope for API call to user.read
        //private string[] scopes = new string[] { "Group.Read.All", "Group.ReadWrite.All", "profile", "User.Read" };

        // The MSAL Public client GraphServiceClient
        private static IPublicClientApplication PublicClientApp;

       // private static string MSGraphURL = "https://graph.microsoft.com/v1.0/";
        private static AuthenticationResult authResult;


        private StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;

        private static Serilog.Core.Logger logger;

        //string redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri().ToString();
        // ms-app://s-1-15-2-148375016-475961868-2312470711-1599034693-979352800-1769312473-2847594358/

        GraphServiceClient graphClient;
        public MainPage()
        {
            CreateLogger();
            this.InitializeComponent();
            lblMessage.Text = config.Tenant;
        }

        public void CreateLogger()
        {
            //var path = ApplicationData.Current.LocalFolder.Path;
            var path = storageFolder.Path;
            var logger = new LoggerConfiguration()
                .WriteTo.File(path + @"\log.txt", 
                rollingInterval: RollingInterval.Hour, 
                rollOnFileSizeLimit: true).CreateLogger();
            logger.Information("Start MetaPlanner");
        }
            

        private async void CleanSharepointList(string listName)
        {
            var items = await graphClient.Sites[config.Site].Lists[listName].Items.Request().GetAsync();
            RadDataGrid.DataContext = items;
            List<ListItem> allItems = new List<ListItem>();
            while (items.Count > 0)
            {
                allItems.AddRange(items);
                if (items.NextPageRequest != null)
                {
                    items = await items.NextPageRequest.GetAsync();
                }
                else
                {
                    break;
                }
            }
            foreach (ListItem item in allItems)
            {
                try
                {
                    await graphClient.Sites[config.Site].Lists[listName].Items[item.Id].Request().DeleteAsync();
                }
                catch(Exception ex)
                {
                    logger.Error(ex.Message);
                }
            }

            /*
            StringBuilder sbDelete = new StringBuilder("<Batch>");
            for (int x = allItems.Count - 1; x >= 0; x--)
            {

                sbDelete.Append("<Method>");
                sbDelete.Append("<SetList Scope='Request'>" + allItems[x].Id.ToString() + "</SetList>");
                sbDelete.Append("<SetVar Name='Cmd'>DELETE</SetVar>");
                sbDelete.Append("<SetVar Name='ID'>listItems[x].ID</SetVar>");
                sbDelete.Append("</Method>");
            }
            sbDelete.Append("</Batch>");
            web.AllowUnsafeUpdates = True;
            web.ProcessBatchData(sbDelete.ToString());*/
        }

        private  void CleanAllSharePointLists()
        {
            CleanSharepointList("assignees");
            CleanSharepointList("tasks");
            CleanSharepointList("buckets");
            //CleanSharepointList("plans");


        }


        private async Task LoadData()
        {

            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                graphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                // Call the /me endpoint of Graph
                User graphUser = await graphClient.Me.Request().GetAsync();


                // Call of Graph

                /*var groups = await graphClient.Groups.Request().GetAsync();
                PlanGrid.DataContext = groups;


                var site = await graphClient.Sites[config.Site].Request().GetAsync();
                PlanGrid.DataContext = site;

                var lists = await graphClient.Sites[config.Site].Lists.Request().GetAsync();
                PlanGrid.DataContext = lists;*/


                var list = await graphClient.Sites[config.Site].Lists["plans"].Request().GetAsync();
                RadDataGrid.DataContext = list;


                // Go back to the UI thread to make changes to the UI
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                {
                    DisplayBasicTokenInfo(authResult);
                    this.SignOutButton.Visibility = Visibility.Visible;
                });
            }
            catch (MsalException msalEx)
            {
                await DisplayMessageAsync($"Error Acquiring Token:{System.Environment.NewLine}{msalEx}");
                logger.Error(msalEx.Message);
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                logger.Error(ex.Message);
                return;
            }
        }


        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGroupButton_Click(object sender, RoutedEventArgs e)
        {
            
            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
            RadDataGrid.StartBringIntoView();


            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                graphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                //var users = await graphClient.Users.Request().GetAsync();

                CleanAllSharePointLists();

                var plans = await graphClient.Me.Planner.Plans.Request().GetAsync();

                List<MetaPlannerPlan> listPlan = new List<MetaPlannerPlan>();
                List<MetaPlannerBucket> listBuckets = new List<MetaPlannerBucket>();
                List<MetaPlannerTask> listTasks = new List<MetaPlannerTask>();
                List<MetaPlannerAssignment> listAssignment = new List<MetaPlannerAssignment>();

                List<PlannerPlan> allPlans = new List<PlannerPlan>();
                RadDataGrid.DataContext = listPlan;

                while (plans.Count > 0)
                {
                    allPlans.AddRange(plans);
                    if (plans.NextPageRequest != null)
                    {
                        plans = await plans.NextPageRequest.GetAsync();
                    }
                    else
                    {
                        break;
                    }
                }


                this.RadialGauge.MaxValue = allPlans.Count;
                this.RadialGauge.TickStep = allPlans.Count/12;
                this.RadialGauge.LabelStep = allPlans.Count/4;
                this.RadialGauge.Visibility = Visibility.Visible;


                lblMessage.Text = "All:" + allPlans.Count;
                int counter = 0;
                foreach (PlannerPlan p in allPlans)
                {
                    var group = await graphClient.Groups[p.Owner].Request().GetAsync();

                    listPlan.Add(new MetaPlannerPlan()
                    {
                        PlanId = p.Id,
                        PlanName = p.Title,
                        CreatedBy = p.CreatedBy.User.Id,
                        CreatedDate = p.CreatedDateTime.ToString(),
                        GroupName = group.DisplayName,
                        GroupDescription = group.Description,
                        GroupMail = group.Mail,
                        Url = "https://tasks.office.com/congenrep.onmicrosoft.com/Home/PlanViews/"+p.Id
                    });

                    

                    var planItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", p.Id},
                                {"PlanName", p.Title},
                                {"CreatedBy", p.CreatedBy.User.Id},
                                {"CreatedDate",  p.CreatedDateTime},
                                {"GroupName",  group.DisplayName},
                                {"GroupDescription",  group.Description},
                                {"GroupMail",  group.Mail},
                                {"Url", "https://tasks.office.com/congenrep.onmicrosoft.com/Home/PlanViews/"+p.Id}
                            }
                        }
                    };
                    await graphClient.Sites[config.Site].Lists["plans"].Items.Request().AddAsync(planItem);

                    counter++;


                    var buckets = await graphClient.Planner.Plans[p.Id].Buckets.Request().GetAsync();

                    List<PlannerBucket> allBuckets = new List<PlannerBucket>();
                    while (buckets.Count > 0)
                    {
                        allBuckets.AddRange(buckets);
                        if (plans.NextPageRequest != null)
                        {
                            buckets = await buckets.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    foreach (PlannerBucket b in allBuckets)
                    {
                        listBuckets.Add(new MetaPlannerBucket()
                        {
                            BucketId = b.Id,
                            BucketName = b.Name,
                            OrderHint = b.OrderHint,
                            PlanId = b.PlanId
                        });

                        var bucketItem = new ListItem
                        {
                            Fields = new FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object>()
                                    {
                                        {"Title", b.Id},
                                        {"BucketName", b.Name},
                                        {"OrderHint",  b.OrderHint},
                                        {"PlanId", b.PlanId}
                                    }
                                }
                        };
                        await graphClient.Sites[config.Site].Lists["buckets"].Items.Request().AddAsync(bucketItem);

                    }
                    var pTasks = await graphClient.Planner.Plans[p.Id].Tasks.Request().GetAsync();


                    List<PlannerTask> allTasks = new List<PlannerTask>();
                    while (pTasks.Count > 0)
                    {
                        allTasks.AddRange(pTasks);
                        if (pTasks.NextPageRequest != null)
                        {
                            pTasks = await pTasks.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    int counterT = 0;
                    foreach (PlannerTask t in allTasks)
                    {
                        MetaPlannerTask myTask = new MetaPlannerTask(){ TaskId= t.Id, Hours="0" };

                        int j = t.Title.IndexOf(";");
                        if (j == -1)
                        {
                            myTask.TaskName = t.Title.Trim();
                        }
                        else
                        {
                            myTask.Prefix = t.Title.Substring(0, j).Trim().ToUpper();

                            string two = t.Title.Substring(j + 1).Trim();
                            int k = two.IndexOf(";");
                            if (k == -1)
                            {
                                myTask.TaskName = two.Trim();
                            }
                            else
                            {
                                myTask.Hours = two.Substring(0, k).Trim();
                                myTask.TaskName = two.Substring(k + 1).Trim();
                            }
                        }

                        #region TaskBody
                        myTask.ActiveChecklistItemCount = t.ActiveChecklistItemCount.ToString();// TODO
                        myTask.AdditionalData = t.AdditionalData.Count.ToString();// TODO
                        myTask.Category1 = t.AppliedCategories.Category1.ToString(); //TODO Make table?
                        myTask.Category2 = t.AppliedCategories.Category2.ToString();
                        myTask.Category3 = t.AppliedCategories.Category3.ToString();
                        myTask.Category4 = t.AppliedCategories.Category4.ToString();
                        myTask.Category5 = t.AppliedCategories.Category5.ToString();
                        myTask.Category6 = t.AppliedCategories.Category6.ToString();
                        myTask.AssigneePriority = t.AssigneePriority;
                        myTask.AssignmentsCount = t.Assignments.Count.ToString();
                        myTask.BucketId = t.BucketId;
                        myTask.ChecklistItemCount = t.ChecklistItemCount.ToString();
                        if (t.CompletedBy != null)
                            myTask.CompletedBy = t.CompletedBy.User.Id;
                        myTask.CompletedDateTime = t.CompletedDateTime.ToString();
                        myTask.ConversationThreadId = t.ConversationThreadId;
                        myTask.CreatedBy = t.CreatedBy.User.Id;
                        myTask.CreatedDateTime = t.CreatedDateTime.ToString();
                        myTask.DueDateTime = t.DueDateTime.ToString();
                        myTask.HasDescription = t.HasDescription.ToString();
                        myTask.OrderHint = t.OrderHint;
                        myTask.PercentComplete = t.PercentComplete.ToString();
                        myTask.PlanId = t.PlanId;
                        myTask.ReferenceCount = t.ReferenceCount.ToString();
                        myTask.StartDateTime = t.StartDateTime.ToString();
                        myTask.Url = "https://tasks.office.com/congenrep.onmicrosoft.com/es-es/Home/Task/" + t.Id;
                        #endregion

                        var taskItem = new ListItem
                        {
                            Fields = new FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object>()
                                    {
                                        {"Title", myTask.TaskId},
                                        {"TaskName", myTask.TaskName},
                                        {"Prefix", myTask.Prefix},

                                        {"Hours", Convert.ToDecimal(myTask.Hours) },

                                        {"ActiveChecklistItemCount", t.ActiveChecklistItemCount},
                                        {"AdditionalData",  t.AdditionalData.Count},
                                        {"Category1", myTask.Category1},
                                        {"Category2", myTask.Category2},
                                        {"Category3", myTask.Category3},
                                        {"Category4", myTask.Category4},
                                        {"Category5", myTask.Category5},
                                        {"Category6", myTask.Category6},
                                        {"AssigneePriority", myTask.AssigneePriority},
                                        {"AssignmentsCount", t.Assignments.Count},
                                        {"BucketId", myTask.BucketId},
                                        {"ChecklistItemCount", t.ChecklistItemCount},
                                        {"CompletedBy", myTask.CompletedBy},
                                        {"CompletedDateTime", t.CompletedDateTime},
                                        {"ConversationThreadId", myTask.ConversationThreadId},
                                        {"CreatedBy", myTask.CreatedBy},
                                        {"CreatedDateTime", t.CreatedDateTime},
                                        {"DueDateTime", t.DueDateTime},
                                        {"HasDescription", myTask.HasDescription},
                                        {"OrderHint", myTask.OrderHint},
                                        {"PercentComplete", t.PercentComplete},
                                        {"PlanId", t.PlanId},
                                        {"ReferenceCount", t.ReferenceCount},
                                        {"StartDateTime", t.StartDateTime},
                                        {"Url", myTask.Url}
                                    }
                            }
                        };
                        await graphClient.Sites[config.Site].Lists["tasks"].Items.Request().AddAsync(taskItem);

                        counterT++;


                        /*
                        TimeSpan ts = TimeSpan.FromMilliseconds(50);
                        var plannerTaskDetails = await graphClient.Planner.Tasks[t.Id].Details.Request().GetAsync();
                        if (plannerTaskDetails.Description != null && plannerTaskDetails.Description.Trim().Length > 0) { 
                            myTask.Details = plannerTaskDetails.Description;
                            TimeTracker.Text = counterT + " " +DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
                            }*/


                        listTasks.Add(myTask);

                        foreach (string userId in t.Assignments.Assignees)
                        {
                            listAssignment.Add(new MetaPlannerAssignment()
                            {
                                TaskId = t.Id,
                                UserId = userId
                            });


                            var assigneesItem = new ListItem
                            {
                                Fields = new FieldValueSet
                                {
                                    AdditionalData = new Dictionary<string, object>()
                                    {
                                        {"Title", t.Id},
                                        {"UserId", userId}
                                    }
                                }
                            };
                            await graphClient.Sites[config.Site].Lists["assignees"].Items.Request().AddAsync(assigneesItem);

                        }

                        lblMessage.Text = counter + " of " + allPlans.Count;
                        Bar.Value = counter;
                    }


                    RadDataGrid.DataContext = listPlan;
                    RadDataGrid.UpdateLayout();
                }


                String prefix = String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second);
                

                Writer writer = new Writer();
                //writer.Write(listPlan, storageFolder, "plans.csv");
                writer.Write(listPlan, storageFolder, prefix+" plans.csv");

                //writer.Write(listBuckets, storageFolder, "buckets.csv");
                writer.Write(listBuckets, storageFolder, prefix + " buckets.csv");

                //writer.Write(listTasks, storageFolder, "tasks.csv");
                writer.Write(listTasks, storageFolder, prefix + " tasks.csv");

                //writer.Write(listAssignment, storageFolder, "assignees.csv");
                writer.Write(listAssignment, storageFolder, prefix + " assignees.csv");

                //File.Copy(sourceFile, destinationFile, true);
                 
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);

            }
            catch (MsalException msalEx)
            {
                await DisplayMessageAsync($"Error Acquiring Token:{System.Environment.NewLine}{msalEx}");
                logger.Error(msalEx.Message);
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                logger.Error(ex.Message);
                return;
            }
        }
        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadData();
        }

        /// <summary>
        /// Signs in the user and obtains an Access token for MS Graph
        /// </summary>
        /// <param name="scopes"></param>
        /// <returns> Access Token</returns>
        private static async System.Threading.Tasks.Task<string> SignInUserAndGetTokenUsingMSAL(string[] scopes)
        {
            // Initialize the MSAL library by building a public client application

            /*
            PublicClientApp = PublicClientApplicationBuilder.Create(config.ClientId)
                .WithAuthority(config.Authority)
                .WithUseCorporateNetwork(false)
                .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                 .WithLogging((level, message, containsPii) =>
                 {
                     Debug.WriteLine($"MSAL: {level} {message} ");
                 }, LogLevel.Warning, enablePiiLogging: false, enableDefaultPlatformLogging: true)
                .Build();
            */

            PublicClientApp = PublicClientApplicationBuilder.Create(config.ClientId)
                .WithAuthority("https://login.microsoftonline.com/common")
                .WithUseCorporateNetwork(false)
                .WithDefaultRedirectUri()
                .WithLogging((level, message, containsPii) =>
                {
                    Debug.WriteLine($"MSAL: {level} {message} ");
                }, LogLevel.Warning, enablePiiLogging: false, enableDefaultPlatformLogging: true)
                .Build();

            // It's good practice to not do work on the UI thread, so use ConfigureAwait(false) whenever possible.
            IEnumerable<IAccount> accounts = await PublicClientApp.GetAccountsAsync().ConfigureAwait(false);
            IAccount firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await PublicClientApp.AcquireTokenSilent(scopes, firstAccount)
                                                  .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync. This indicates you need to call AcquireTokenAsync to acquire a token
                Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");
                logger.Error(ex.Message);
                authResult = await PublicClientApp.AcquireTokenInteractive(scopes)
                                                  .ExecuteAsync()
                                                  .ConfigureAwait(false);

            }
            return authResult.AccessToken;
        }

        /// <summary>
        /// Sign in user using MSAL and obtain a token for Microsoft Graph
        /// </summary>
        /// <returns>GraphServiceClient</returns>
        private async static Task<GraphServiceClient> SignInAndInitializeGraphServiceClient(string[] scopes)
        {
            GraphServiceClient graphClient = new GraphServiceClient(config.MSGraphURL,
                new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await SignInUserAndGetTokenUsingMSAL(scopes));
                }));

            return await Task.FromResult(graphClient);
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
             IEnumerable<IAccount> accounts = await PublicClientApp.GetAccountsAsync().ConfigureAwait(false);
             IAccount firstAccount = accounts.FirstOrDefault();

             try
             {
                 await PublicClientApp.RemoveAsync(firstAccount).ConfigureAwait(false);
                 await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                 {

                     this.btnCall.Visibility = Visibility.Visible;
                     this.SignOutButton.Visibility = Visibility.Collapsed;
                 });
             }
             catch (MsalException ex)
             {
                logger.Error(ex.Message);
             }    
         }

        /// <summary>
        /// Display basic information contained in the token. Needs to be called from the UI thead.
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {

            if (authResult != null)
            {
                lblMessage.Text = authResult.Account.Username;
            }
        }

        /// <summary>
        /// Displays a message in the ResultText. Can be called from any thread.
        /// </summary>
        private async Task DisplayMessageAsync(string message)
        {
            await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
                   () =>
                   {
                       lblMessage.Text = message;      
                    });

            // Create the message dialog and set its content
            var messageDialog = new MessageDialog(message,"Error");

            // Set the command that will be invoked by default
           // messageDialog.DefaultCommandIndex = 0;

            // Set the command to be invoked when escape is pressed
           // messageDialog.CancelCommandIndex = 1;

            // Show the message dialog
            await messageDialog.ShowAsync();
        }





    }
}
