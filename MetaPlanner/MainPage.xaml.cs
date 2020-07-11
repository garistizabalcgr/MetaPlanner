using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using CsvHelper;
using System.Globalization;
using MetaPlanner.Model;
using MetaPlanner.Output;
using Telerik.Core;

// La plantilla de elemento Página en blanco está documentada en https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0xc0a

namespace MetaPlanner
{
    /// <summary>
    /// Página vacía que se puede usar de forma independiente o a la que se puede navegar dentro de un objeto Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        //Set the scope for API call to user.read
        private string[] scopes = new string[] { "Group.Read.All", "Group.ReadWrite.All", "profile", "User.Read" };

        //(Application Id) of your app registration and the tenant information. 
        private const string ClientId = "095ada9d-71d5-42b3-a962-28726a951818";

        private const string Tenant = "congenrep.onmicrosoft.com"; // Alternatively "[Enter your tenant, as obtained from the azure portal, e.g. kko365.onmicrosoft.com]"
        private const string Authority = "https://login.microsoftonline.com/" + Tenant;

        // The MSAL Public client app
        private static IPublicClientApplication PublicClientApp;

        private static string MSGraphURL = "https://graph.microsoft.com/v1.0/";
        private static AuthenticationResult authResult;

        //string redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri().ToString();
        // ms-app://s-1-15-2-148375016-475961868-2312470711-1599034693-979352800-1769312473-2847594358/

        public MainPage()
        {
            this.InitializeComponent();
            var task = Task.Run(async () => await LoadData());

        }

        private async Task LoadData()
        {

            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphServiceClient graphClient = await SignInAndInitializeGraphServiceClient(scopes);

                // Call the /me endpoint of Graph
                User graphUser = await graphClient.Me.Request().GetAsync();

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
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                return;
            }
        }


        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGroupButton_Click(object sender, RoutedEventArgs e)
        {
            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphServiceClient graphClient = await SignInAndInitializeGraphServiceClient(scopes);

                // Call of Graph
                // var tasks = await graphClient.Me.Planner.Tasks.Request().GetAsync();


                //var users = await graphClient.Users.Request().GetAsync();

                Windows.Storage.StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;

                var plans = await graphClient.Me.Planner.Plans.Request().GetAsync();

                List<Plan> listPlan = new List<Plan>();
                List<Bucket> listBuckets = new List<Bucket>();
                List<PTask> listTasks = new List<PTask>();
                List<Assignment> listAssignment = new List<Assignment>();

                List<PlannerPlan> allPlans = new List<PlannerPlan>();
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

                lblMessage.Text = "All:" + allPlans.Count;
                int counter = 0;
                foreach (PlannerPlan p in allPlans)
                {
                    listPlan.Add(new Plan()
                    {
                        PlanId = p.Id,
                        PlanName = p.Title,
                        CreatedBy = p.CreatedBy.User.Id,
                        CreatedDate = p.CreatedDateTime.ToString(),
                        Owner = p.Owner
                    });
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
                        listBuckets.Add(new Bucket()
                        {
                            BucketId = b.Id,
                            BucketName = b.Name,
                            OrderHint = b.OrderHint,
                            PlanId = b.PlanId
                        });
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
                        PTask myTask = new PTask();
                        myTask.TaskId = t.Id;
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
                        myTask.ChecklistItemCount = t.ReferenceCount.ToString();
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
                            listAssignment.Add(new Assignment()
                            {
                                TaskId = t.Id,
                                UserId = userId
                            });
                        }

                        lblMessage.Text = counter + " of " + allPlans.Count;

                    }


                }


                Writer writer = new Writer();
                writer.WritePlans(listPlan, "plans");
                writer.WriteBuckets(listBuckets, "buckets");
                writer.WriteTasks(listTasks, "tasks");
                writer.WriteAssignees(listAssignment, "assignees");


                PlanManager.SetPlans(listPlan);
                this.DataContext = PlanManager.GetPlans();
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);

            }
            catch (MsalException msalEx)
            {
                await DisplayMessageAsync($"Error Acquiring Token:{System.Environment.NewLine}{msalEx}");
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
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
            PublicClientApp = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(Authority)
                .WithUseCorporateNetwork(false)
                .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                 .WithLogging((level, message, containsPii) =>
                 {
                     Debug.WriteLine($"MSAL: {level} {message} ");
                 }, LogLevel.Warning, enablePiiLogging: false, enableDefaultPlatformLogging: true)
                .Build();
            */

            PublicClientApp = PublicClientApplicationBuilder.Create(ClientId)
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
            GraphServiceClient graphClient = new GraphServiceClient(MSGraphURL,
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
                lblMessage.Text = ex.Message;
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
        }

        public class Foo
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }
    }
}
