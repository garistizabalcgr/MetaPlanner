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
using System.Xml.Serialization;
using Windows.UI.Xaml.Data;
using System.Collections.Immutable;



// La plantilla de elemento Página en blanco está documentada en https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0xc0a

namespace MetaPlanner
{
    /// <summary>
    /// Página vacía que se puede usar de forma independiente o a la que se puede navegar dentro de un objeto Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        // Configuration
        public static AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile();

        // The MSAL Public client GraphServiceClient
        private static IPublicClientApplication PublicClientApp;

        // private static string MSGraphURL = "https://graph.microsoft.com/v1.0/";
        private static AuthenticationResult authResult;

        //Folder of file storage
        private StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;

        //Client to Cloud Service
        private GraphServiceClient GraphClient;

        //Dictionary of Plans
        private Dictionary<string, MetaPlannerPlan> PlannerPlans = new Dictionary<string, MetaPlannerPlan>();

        //Dictionary of Buckets
        private Dictionary<string, MetaPlannerBucket> PlannerBuckets = new Dictionary<string, MetaPlannerBucket>();

        //Dictionary of Tasks
        private Dictionary<string, MetaPlannerTask> PlannerTasks = new Dictionary<string, MetaPlannerTask>();

        //Dictionary of Assignments
        private Dictionary<string, MetaPlannerAssignment> PlannerAssignments = new Dictionary<string, MetaPlannerAssignment>();

        //Dictionary of Users
        private Dictionary<string, MetaPlannerUser> PlannerUsers = new Dictionary<string, MetaPlannerUser>();

        //Date of creation
        private String TimeStamp
        {
            get
            {
                return String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second);
            }
        }
        //Cvs writer of files}
        private Writer writer = new Writer();

        //string redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri().ToString();
        // ms-app://s-1-15-2-148375016-475961868-2312470711-1599034693-979352800-1769312473-2847594358/



        public MainPage()
        {
           
            this.InitializeComponent();
            lblMessage.Text = config.Tenant;
        }
        


        private async Task<List<ListItem>> GetSharePointList(string listName)
        {
            var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("expand", "fields")
                };
            var items = await GraphClient.Sites[config.Site].Lists[listName].Items.Request(queryOptions).GetAsync();
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
            return allItems;
        }

        private async Task CleanSharepointList(string listName)
        {
            var items = await GraphClient.Sites[config.Site].Lists[listName].Items.Request().GetAsync();
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
                    await GraphClient.Sites[config.Site].Lists[listName].Items[item.Id].Request().DeleteAsync();
                }
                catch (Exception ex)
                {
                    App.logger.Error(ex.Message);
                }
            }
        }

        private async void CleanAllSharePointLists(object sender, RoutedEventArgs e)
        {
            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

            GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

            await CleanSharepointList("users");
            await CleanSharepointList("assignees");
            await CleanSharepointList("tasks");
            await CleanSharepointList("buckets");
            await CleanSharepointList("plans");

            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
        }

        private async Task LoadData()
        {
            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                // Call the /me endpoint of Graph
                User graphUser = await GraphClient.Me.Request().GetAsync();


                // Call of Graph

                /*var groups = await graphClient.Groups.Request().GetAsync();
                PlanGrid.DataContext = groups;


                var site = await graphClient.Sites[config.Site].Request().GetAsync();
                PlanGrid.DataContext = site;

                var lists = await graphClient.Sites[config.Site].Lists.Request().GetAsync();
                PlanGrid.DataContext = lists;*/


                var list = await GraphClient.Sites[config.Site].Lists["plans"].Request().GetAsync();
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
                App.logger.Error(msalEx.Message);
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
                return;
            }
        }

        /// <summary>
        /// Pattern of Call Commando interactive - Description
        /// </summary>
        private async void Pattern_Command(object sender, RoutedEventArgs e) {
            try
            {
                App.logger.Information("Start Command");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                //TODO: Complete Code


                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End Command");
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
            }
        }

        private async Task WriteAndUpload(System.Collections.IDictionary dictionary, string name)
        {
            string fileName = name + ".csv";
            await writer.Write(dictionary.Values, storageFolder, fileName);

            FileStream fs = new FileStream(storageFolder.Path + "\\" + fileName, FileMode.Open, FileAccess.Read);
           
            DriveItem driveItem = new DriveItem();
            driveItem.Name = fileName;
            driveItem.File = new Microsoft.Graph.File();

            var drive = await GraphClient.Sites[config.Site].Drive.Request().GetAsync();
            try
            {
                var file = await GraphClient.Sites[config.Site].Drive.Root.Children[driveItem.Name].Request().GetAsync();
                var resOld = await GraphClient.Sites[config.Site].Drive.Items[file.Id].Content.Request().PutAsync<DriveItem>(fs);
            }
            catch(Exception ex)
            {
                driveItem = await GraphClient.Sites[config.Site].Drive.Root.Children.Request().AddAsync(driveItem);
                var resNew = await GraphClient.Sites[config.Site].Drive.Items[driveItem.Id].Content.Request().PutAsync<DriveItem>(fs);
                lblMessage.Text = ex.Message;
            }

            //TimeStamp to versioning an historic trace
            if (!TimeStamp.Trim().Equals(""))
            {
                string fileNameT = TimeStamp + " " + name + ".csv";
                await writer.Write(dictionary.Values, storageFolder, fileNameT);
                FileStream fsT = new FileStream(storageFolder.Path + "\\" + fileNameT, FileMode.Open, FileAccess.Read);
                DriveItem driveItemStamp = new DriveItem();
                driveItemStamp.Name = fileNameT;
                driveItemStamp.File = new Microsoft.Graph.File();
                try
                {
                    var fileT = await GraphClient.Sites[config.Site].Drive.Root.Children[driveItemStamp.Name].Request().GetAsync();
                    var resOldT = await GraphClient.Sites[config.Site].Drive.Items[fileT.Id].Content.Request().PutAsync<DriveItem>(fsT);
                }
                catch (Exception ex1)
                {
                    driveItemStamp = await GraphClient.Sites[config.Site].Drive.Root.Children.Request().AddAsync(driveItemStamp);
                    var resNewT = await GraphClient.Sites[config.Site].Drive.Items[driveItemStamp.Id].Content.Request().PutAsync<DriveItem>(fsT);
                    lblMessage.Text = ex1.Message;
                }
            }
        }

        #region Plans

        /// <summary>
        /// Get all data from Plans from Planner in plannerPlans
        /// </summary>
        private async Task GetPlannerPlans()
        {
            var page = await GraphClient.Me.Planner.Plans.Request().GetAsync();
            List<PlannerPlan> listPlanner = new List<PlannerPlan>();
            while (page.Count > 0)
            {
                listPlanner.AddRange(page);
                if (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync();
                }
                else
                {
                    break;
                }
            }
            PlannerPlans = new Dictionary<string, MetaPlannerPlan>();
            foreach (PlannerPlan p in listPlanner)
            {
                var group = await GraphClient.Groups[p.Owner].Request().GetAsync();
                PlannerPlans.Add(p.Id,
                    new MetaPlannerPlan()
                    {
                        PlanId = p.Id,
                        PlanName = p.Title,
                        CreatedBy = p.CreatedBy.User.Id,
                        CreatedDate = p.CreatedDateTime,
                        GroupName = group.DisplayName,
                        GroupDescription = group.Description,
                        GroupMail = group.Mail,
                        Url = "https://tasks.office.com/" + config.Tenant + "/Home/PlanViews/" + p.Id
                    });
            }
        }
        
        /// <summary>
        /// Process Plans Data
        /// </summary>
        private async void ProcessPlans(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessPlans");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

               await GetPlannerPlans();

                await WriteAndUpload(PlannerPlans, "plans");

                if (config.IsSharePointListEnabled.Equals("true"))
                {
                    #region Get bulk data from SharePoint
                    var listPlans = await GetSharePointList("plans");

                    Dictionary<string, MetaPlannerPlan> sharePointPlans = new Dictionary<string, MetaPlannerPlan>();
                    Dictionary<string, string> itemIds = new Dictionary<string, string>();
                    Dictionary<string, ListItem> items = new Dictionary<string, ListItem>();

                    foreach (ListItem item in listPlans)
                    {
                        MetaPlannerPlan plan = new MetaPlannerPlan(item.Fields.AdditionalData);
                        sharePointPlans.Add(plan.PlanId, plan);
                        itemIds.Add(plan.PlanId, item.Id);
                        items.Add(item.Id, item);
                    }

                    #endregion
                    await ConciliationPlans(sharePointPlans, itemIds, items);
                }

                RadDataGrid.DataContext = PlannerPlans.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessPlans");
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"ProcessPlans:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
                return;
            }
        }

        private async Task ConciliationPlans(Dictionary<string, MetaPlannerPlan> sharePointPlans, Dictionary<string, string> itemIds, Dictionary<string, ListItem> items)
        {
            int add = 0;
            int del = 0;
            int upd = 0;

            #region Add from planner not in sharepoint
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerPlan> entry in PlannerPlans)
            {
                if ( ! sharePointPlans.ContainsKey(entry.Key))
                {
                    var planItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", entry.Value.PlanId},
                                {"PlanName", entry.Value.PlanName},
                                {"CreatedBy", entry.Value.CreatedBy},
                                {"CreatedDate", entry.Value.CreatedDate },
                                {"GroupName",  entry.Value.GroupName },
                                {"GroupDescription",  entry.Value.GroupDescription},
                                {"GroupMail",  entry.Value.GroupMail},
                                {"Url", entry.Value.Url}
                            }
                        }
                    };
                    await GraphClient.Sites[config.Site].Lists["plans"].Items.Request().AddAsync(planItem);
                    add++;
                    lblMessage.Text = "Plan A: " + add + " D: " + del + " U:" + upd;
                }               
            }
            #endregion

            #region Delete in sharepoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerPlan> entry in sharePointPlans)
            {
                if ( ! PlannerPlans.ContainsKey(entry.Key))
                {
                    await GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    del++;
                    lblMessage.Text = "Plan A: " + add + " D: " + del + " U:" + upd;
                }
            }
            #endregion

            #region Update in Sharepoint changes from planner
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerPlan> entry in PlannerPlans)
            {
                if (sharePointPlans.ContainsKey(entry.Key))
                {
                    MetaPlannerPlan origin = PlannerPlans[entry.Key];
                    MetaPlannerPlan destination = sharePointPlans[entry.Key];

                    Dictionary<string, object> additionalData = new Dictionary<string, object>();
                    if (!String.Equals(origin.PlanName, destination.PlanName))
                    {
                        additionalData.Add("PlanName", origin.PlanName);
                    }
                    if (! String.Equals(origin.GroupDescription, destination.GroupDescription))
                    {
                        additionalData.Add("GroupDescription", origin.GroupDescription);
                    }
                    if (!String.Equals(origin.GroupMail, destination.GroupMail))
                    {
                        additionalData.Add("GroupMail", origin.GroupMail);
                    }
                    if (!String.Equals(origin.GroupName, destination.GroupName))
                    {
                        additionalData.Add("GroupName", origin.GroupName);
                    }

                    if (additionalData.Keys.Count > 0)
                    {
                        FieldValueSet fieldsChange = new FieldValueSet();
                        fieldsChange.AdditionalData = additionalData;
                        await GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        upd++;
                        lblMessage.Text = "Plan A: " + add + " D: " + del + " U:" + upd;
                    }
                }
            }
            #endregion
            App.logger.Information("Plan Added: " + add + " Deleted: " + del + " Updated:" + upd);
        }
        #endregion


        #region Buckets

        /// <summary>
        /// Get all data from Plans from Planner in plannerPlans
        /// </summary>
        private async Task GetPlannerBuckets()
        {
            if (PlannerPlans == null || PlannerPlans.Count == 0)
            {
                await GetPlannerPlans();
            }

            PlannerBuckets = new Dictionary<string, MetaPlannerBucket>();
            foreach (MetaPlannerPlan plan in PlannerPlans.Values)
            {
                var buckets = await GraphClient.Planner.Plans[plan.PlanId].Buckets.Request().GetAsync();
                List<PlannerBucket> listBuckets = new List<PlannerBucket>();
                while (buckets.Count > 0)
                {
                    listBuckets.AddRange(buckets);
                    if (buckets.NextPageRequest != null)
                    {
                        buckets = await buckets.NextPageRequest.GetAsync();
                    }
                    else
                    {
                        break;
                    }
                }
               
                foreach (PlannerBucket bucket in listBuckets)
                {
                    PlannerBuckets.Add(bucket.Id,new MetaPlannerBucket()
                    {
                        BucketId = bucket.Id,
                        BucketName = bucket.Name,
                        OrderHint = bucket.OrderHint,
                        PlanId = plan.PlanId
                    });
                }
            }
        }

        /// <summary
        /// Process Buckets Data
        /// </summary>
        private async void ProcessBuckets(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessBuckets");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                await GetPlannerBuckets();
                await WriteAndUpload(PlannerBuckets, "buckets");

                if (config.IsSharePointListEnabled.Equals("true"))
                {
                    #region Get bulk data from SharePoint
                    var listBuckets = await GetSharePointList("buckets");

                    Dictionary<string, MetaPlannerBucket> sharePointBuckets = new Dictionary<string, MetaPlannerBucket>();
                    Dictionary<string, string> itemIds = new Dictionary<string, string>();
                    Dictionary<string, ListItem> items = new Dictionary<string, ListItem>();

                    foreach (ListItem item in listBuckets)
                    {
                        MetaPlannerBucket bucket = new MetaPlannerBucket(item.Fields.AdditionalData);
                        sharePointBuckets.Add(bucket.BucketId, bucket);
                        itemIds.Add(bucket.BucketId, item.Id);
                        items.Add(item.Id, item);
                    }
                    #endregion

                    await ConciliationBuckets(sharePointBuckets, itemIds, items);
                }

                RadDataGrid.DataContext = PlannerBuckets.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessBuckets");
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"ProcessBuckets:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
                return;
            }
        }

        private async Task ConciliationBuckets(Dictionary<string, MetaPlannerBucket> sharePointBuckets, Dictionary<string, string> itemIds, Dictionary<string, ListItem> items)
        {
            int add = 0;
            int del = 0;
            int upd = 0;

            #region Add from planner not in sharepoint
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerBucket> entry in PlannerBuckets)
            {
                if (!sharePointBuckets.ContainsKey(entry.Key))
                {
                    var bucketItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", entry.Value.BucketId},
                                {"BucketName", entry.Value.BucketName},
                                {"PlanId", entry.Value.PlanId},
                                {"OrderHint", entry.Value.OrderHint }
                            }
                        }
                    };
                    await GraphClient.Sites[config.Site].Lists["buckets"].Items.Request().AddAsync(bucketItem);
                    add++;
                    lblMessage.Text = "Bucket A: " + add + " D: " + del + " U:" + upd;
                }
            }
            #endregion

            #region Delete in SharePoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerBucket> entry in sharePointBuckets)
            {
                if (! PlannerBuckets.ContainsKey(entry.Key))
                {
                    await GraphClient.Sites[config.Site].Lists["buckets"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    del++;
                    lblMessage.Text = "Bucket A: " + add + " D: " + del + " U:" + upd;
                }
            }
            #endregion

            #region Update in Sharepoint changes from planner
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerBucket> entry in PlannerBuckets)
            {
                if (sharePointBuckets.ContainsKey(entry.Key))
                {
                    MetaPlannerBucket origin = PlannerBuckets[entry.Key];
                    MetaPlannerBucket destination = sharePointBuckets[entry.Key];

                    Dictionary<string, object> additionalData = new Dictionary<string, object>();
                    if (!String.Equals(origin.BucketName, destination.BucketName))
                    {
                        additionalData.Add("BucketName", origin.BucketName);
                    }
                    if (!String.Equals(origin.OrderHint, destination.OrderHint))
                    {
                        additionalData.Add("OrderHint", origin.OrderHint);
                    }
                    if (additionalData.Keys.Count > 0)
                    {
                        FieldValueSet fieldsChange = new FieldValueSet();
                        fieldsChange.AdditionalData = additionalData;
                        await GraphClient.Sites[config.Site].Lists["buckets"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        upd++;
                        lblMessage.Text = "Bucket A: " + add + " D: " + del + " U:" + upd;
                    }
                }
            }
            #endregion

            App.logger.Information("Buckets Added: " + add + " Deleted: " + del + " Updated:" + upd);
        }

        #endregion


        #region Tasks

        /// <summary>
        /// Get all data from Plans from Planner in plannerPlans
        /// </summary>
        private async Task GetPlannerTasks()
        {
            if (PlannerPlans == null || PlannerPlans.Count == 0)
            {
                await GetPlannerPlans();
            }

            PlannerTasks = new Dictionary<string, MetaPlannerTask>();
            PlannerAssignments = new Dictionary<string, MetaPlannerAssignment>();
            PlannerUsers = new Dictionary<string, MetaPlannerUser>(); //TODO

            foreach (MetaPlannerPlan plan in PlannerPlans.Values)
            {
                var tasks = await GraphClient.Planner.Plans[plan.PlanId].Tasks.Request().GetAsync();
                List<PlannerTask> listTasks = new List<PlannerTask>();
                while (tasks.Count > 0)
                {
                    listTasks.AddRange(tasks);
                    if (tasks.NextPageRequest != null)
                    {
                        tasks = await tasks.NextPageRequest.GetAsync();
                    }
                    else
                    {
                        break;
                    }
                }

                foreach (PlannerTask task in listTasks)
                {
                    MetaPlannerTask myTask = new MetaPlannerTask() { TaskId = task.Id, Hours = "0" };

                    #region Task custom fields
                    int j = task.Title.IndexOf(";");
                    if (j == -1)
                    {
                        myTask.TaskName = task.Title.Trim();
                    }
                    else
                    {
                        myTask.Prefix = task.Title.Substring(0, j).Trim().ToUpper();

                        string two = task.Title.Substring(j + 1).Trim();
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
                    #endregion

                    #region TaskBody
                    myTask.PlanId = task.PlanId;
                    myTask.ActiveChecklistItemCount = task.ActiveChecklistItemCount.ToString();
                    myTask.AdditionalData = task.AdditionalData.Count.ToString();
                    myTask.Category1 = task.AppliedCategories.Category1.ToString(); 
                    myTask.Category2 = task.AppliedCategories.Category2.ToString();
                    myTask.Category3 = task.AppliedCategories.Category3.ToString();
                    myTask.Category4 = task.AppliedCategories.Category4.ToString();
                    myTask.Category5 = task.AppliedCategories.Category5.ToString();
                    myTask.Category6 = task.AppliedCategories.Category6.ToString();
                    myTask.AssigneePriority = task.AssigneePriority;
                    myTask.AssignmentsCount = task.Assignments.Count.ToString();
                    myTask.BucketId = task.BucketId;
                    myTask.ChecklistItemCount = task.ChecklistItemCount.ToString();
                    if (task.CompletedBy != null)
                        myTask.CompletedBy = task.CompletedBy.User.Id;
                    myTask.CompletedDateTime = task.CompletedDateTime.ToString();
                    myTask.ConversationThreadId = task.ConversationThreadId;
                    myTask.CreatedBy = task.CreatedBy.User.Id;
                    myTask.CreatedDateTime = task.CreatedDateTime.ToString();
                    myTask.DueDateTime = task.DueDateTime.ToString();
                    myTask.HasDescription = task.HasDescription.ToString();
                    myTask.OrderHint = task.OrderHint;
                    myTask.PercentComplete = task.PercentComplete.ToString();
                    myTask.ReferenceCount = task.ReferenceCount.ToString();
                    myTask.StartDateTime = task.StartDateTime.ToString();
                    myTask.Url = "https://tasks.office.com/"+config.Tenant+"/es-es/Home/Task/" + task.Id;
                    #endregion
                    object priority;
                    task.AdditionalData.TryGetValue("priority", out priority);
                    if (priority != null)
                        myTask.Priority = priority.ToString();
                    PlannerTasks.Add(myTask.TaskId, myTask);
                    GetPlannerAssignment(task);
                }
                App.logger.Information("Start ProcessTasks" + PlannerTasks.Count);
            }
        }

        /// <summary
        /// Process Task Data
        /// </summary>
        /// 
        private void GetPlannerAssignment(PlannerTask task)
        {
            foreach (string userId in task.Assignments.Assignees)
            {
                PlannerAssignments.Add(task.Id+"_"+userId,new MetaPlannerAssignment()
                {
                    TaskId = task.Id,
                    UserId = userId
                });
            }
        }

        /// <summary
        /// Process Task Data
        /// </summary>
        private async void ProcessTasks(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessTasks");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                await GetPlannerTasks();

                await WriteAndUpload(PlannerTasks, "tasks");
                await WriteAndUpload(PlannerAssignments, "assignees");

                if (config.IsSharePointListEnabled.Equals("true"))
                {

                    #region Get bulk data from SharePoint Tasks
                    var listTasks = await GetSharePointList("tasks");

                    Dictionary<string, MetaPlannerTask> sharePointTasks = new Dictionary<string, MetaPlannerTask>();
                    Dictionary<string, string> itemIds = new Dictionary<string, string>();
                    Dictionary<string, ListItem> items = new Dictionary<string, ListItem>();

                    foreach (ListItem item in listTasks)
                    {
                        MetaPlannerTask task = new MetaPlannerTask(item.Fields.AdditionalData);
                        sharePointTasks.Add(task.TaskId, task);
                        itemIds.Add(task.TaskId, item.Id);
                        items.Add(item.Id, item);
                    }
                    #endregion
                    await ConciliationTasks(sharePointTasks, itemIds, items);

                    #region Get bulk data from SharePoint
                    var listAssignees = await GetSharePointList("assignees");

                    Dictionary<string, MetaPlannerAssignment> sharePointAsignees = new Dictionary<string, MetaPlannerAssignment>();
                    Dictionary<string, string> itemIdsA = new Dictionary<string, string>();
                    Dictionary<string, ListItem> itemsA = new Dictionary<string, ListItem>();

                    foreach (ListItem theItem in listAssignees)
                    {
                        MetaPlannerAssignment assignment = new MetaPlannerAssignment(theItem.Fields.AdditionalData);
                        sharePointAsignees.Add(assignment.TaskId+"_"+ assignment.UserId, assignment);
                        itemIdsA.Add(assignment.TaskId + "_" + assignment.UserId, theItem.Id);
                        itemsA.Add(theItem.Id, theItem);
                    }
                    #endregion
                    await ConciliationAssignments(sharePointAsignees, itemIdsA, itemsA);
                }


                RadDataGrid.DataContext = PlannerTasks.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessTasks");
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"ProcessTasks:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
                return;
            }
        }

        private async Task ConciliationTasks(Dictionary<string, MetaPlannerTask> sharePointTasks, Dictionary<string, string> itemIds, Dictionary<string, ListItem> items)
        {
            int add = 0;
            int del = 0;
            int upd = 0;

            #region Add from planner not in sharepoint
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerTask> entry in PlannerTasks)
            {
                if (!sharePointTasks.ContainsKey(entry.Key))
                {
                    var taskItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", entry.Value.TaskId},
                                {"TaskName", entry.Value.TaskName},
                                {"Prefix", entry.Value.Prefix},
                                {"Hours", entry.Value.Hours },
                                {"ActiveChecklistItemCount", entry.Value.ActiveChecklistItemCount },
                                {"AdditionalData", entry.Value.AdditionalData },
                                {"Category1", entry.Value.Category1 },
                                {"Category2", entry.Value.Category2 },
                                {"Category3", entry.Value.Category3 },
                                {"Category4", entry.Value.Category4 },
                                {"Category5", entry.Value.Category5 },
                                {"Category6", entry.Value.Category6 },
                                {"AssigneePriority", entry.Value.AssigneePriority },
                                {"AssignmentsCount", entry.Value.AssignmentsCount },
                                {"BucketId", entry.Value.BucketId },
                                {"PlanId", entry.Value.PlanId },
                                {"ChecklistItemCount", entry.Value.ChecklistItemCount },
                                {"CompletedBy", entry.Value.CompletedBy },
                                {"CompletedDateTime", entry.Value.CompletedDateTime },
                                {"ConversationThreadId", entry.Value.ConversationThreadId },
                                {"CreatedBy", entry.Value.CreatedBy },
                                {"CreatedDateTime", entry.Value.CreatedDateTime },
                                {"DueDateTime", entry.Value.DueDateTime },
                                {"HasDescription", entry.Value.HasDescription },
                                {"OrderHint", entry.Value.OrderHint },
                                {"PercentComplete", entry.Value.PercentComplete },
                                {"ReferenceCount", entry.Value.ReferenceCount },
                                {"StartDateTime", entry.Value.StartDateTime },
                                {"Url", entry.Value.Url},
                                {"Priority",entry.Value.Priority }
                            }
                        }
                    };
                    try
                    {
                        var a = await GraphClient.Sites[config.Site].Lists["tasks"].Items.Request().AddAsync(taskItem);
                        add++;
                        lblMessage.Text = "Task A: " + add + " D: " + del + " U:" + upd;
                    }
                    catch (Exception exAdd)
                    {
                        await DisplayMessageAsync($"Error Adding:{System.Environment.NewLine}{exAdd}");
                        App.logger.Error(exAdd.Message);
                    }
                }
            }
            #endregion

            #region Delete in SharePoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerTask> entry in sharePointTasks)
            {
                if (!PlannerTasks.ContainsKey(entry.Key))
                {
                    try
                    {
                        await GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        del++;
                        lblMessage.Text = "Task A: " + add + " D: " + del + " U:" + upd;
                    }
                    catch (Exception exDel)
                    {
                        await DisplayMessageAsync($"Error Deleting:{System.Environment.NewLine}{exDel}");
                        App.logger.Error(exDel.Message);
                    }
                }
            }
            #endregion

            #region Update in Sharepoint changes from planner
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerTask> entry in PlannerTasks)
            {
                if (sharePointTasks.ContainsKey(entry.Key))
                {
                    MetaPlannerTask origin = PlannerTasks[entry.Key];
                    MetaPlannerTask destination = sharePointTasks[entry.Key];

                    Dictionary<string, object> additionalData = new Dictionary<string, object>();
                    if (!String.Equals(origin.TaskName, destination.TaskName))
                    {
                        additionalData.Add("TaskName", origin.TaskName);
                    }
                    #region Changes
                    if (!String.Equals(origin.Prefix, destination.Prefix))
                    {
                        additionalData.Add("Prefix", origin.Prefix);
                    }

                    if (!String.Equals(origin.Hours, destination.Hours))
                    {
                        additionalData.Add("Hours", origin.Hours);
                    }

                    if (!String.Equals(origin.ActiveChecklistItemCount, destination.ActiveChecklistItemCount))
                    {
                        additionalData.Add("ActiveChecklistItemCount", origin.ActiveChecklistItemCount);
                    }

                    if (!String.Equals(origin.AdditionalData, destination.AdditionalData))
                    {
                        additionalData.Add("AdditionalData", origin.AdditionalData);
                    }

                    if (!String.Equals(origin.Category1, destination.Category1))
                    {
                        additionalData.Add("Category1", origin.Category1);
                    }

                    if (!String.Equals(origin.Category2, destination.Category2))
                    {
                        additionalData.Add("Category2", origin.Category2);
                    }

                    if (!String.Equals(origin.Category3, destination.Category3))
                    {
                        additionalData.Add("Category3", origin.Category3);
                    }

                    if (!String.Equals(origin.Category4, destination.Category4))
                    {
                        additionalData.Add("Category4", origin.Category4);
                    }

                    if (!String.Equals(origin.Category5, destination.Category5))
                    {
                        additionalData.Add("Category5", origin.Category5);
                    }

                    if (!String.Equals(origin.Category6, destination.Category6))
                    {
                        additionalData.Add("Category6", origin.Category6);
                    }

                    if (!String.Equals(origin.AssigneePriority, destination.AssigneePriority))
                    {
                        additionalData.Add("AssigneePriority", origin.AssigneePriority);
                    }

                    if (!String.Equals(origin.AssignmentsCount, destination.AssignmentsCount))
                    {
                        additionalData.Add("AssignmentsCount", origin.AssignmentsCount);
                    }

                    if (!String.Equals(origin.BucketId, destination.BucketId))
                    {
                        additionalData.Add("BucketId", origin.BucketId);
                    }

                    if (!String.Equals(origin.PlanId, destination.PlanId))
                    {
                        additionalData.Add("PlanId", origin.PlanId);
                    }

                    if (!String.Equals(origin.ChecklistItemCount, destination.ChecklistItemCount))
                    {
                        additionalData.Add("ChecklistItemCount", origin.ChecklistItemCount);
                    }

                    if (!String.Equals(origin.CompletedBy, destination.CompletedBy))
                    {
                        additionalData.Add("CompletedBy", origin.CompletedBy);
                    }

                    if (!String.Equals(origin.CompletedDateTime, destination.CompletedDateTime))
                    {
                        additionalData.Add("CompletedDateTime", origin.CompletedDateTime);
                    }

                    if (!String.Equals(origin.ConversationThreadId, destination.ConversationThreadId))
                    {
                        additionalData.Add("ConversationThreadId", origin.ConversationThreadId);
                    }

                    if (!String.Equals(origin.CreatedBy, destination.CreatedBy))
                    {
                        additionalData.Add("CreatedBy", origin.CreatedBy);
                    }

                    if (!String.Equals(origin.CreatedDateTime, destination.CreatedDateTime))
                    {
                        additionalData.Add("CreatedDateTime", origin.CreatedDateTime);
                    }

                    if (!String.Equals(origin.DueDateTime, destination.DueDateTime))
                    {
                        additionalData.Add("DueDateTime", origin.DueDateTime);
                    }

                    if (!String.Equals(origin.HasDescription, destination.HasDescription))
                    {
                        additionalData.Add("HasDescription", origin.HasDescription);
                    }

                    if (!String.Equals(origin.OrderHint, destination.OrderHint))
                    {
                        additionalData.Add("OrderHint", origin.OrderHint);
                    }

                    if (!String.Equals(origin.PercentComplete, destination.PercentComplete))
                    {
                        additionalData.Add("PercentComplete", origin.PercentComplete);
                    }

                    if (!String.Equals(origin.PlanId, destination.PlanId))
                    {
                        additionalData.Add("PlanId", origin.PlanId);
                    }

                    if (!String.Equals(origin.ReferenceCount, destination.ReferenceCount))
                    {
                        additionalData.Add("ReferenceCount", origin.ReferenceCount);
                    }

                    if (!String.Equals(origin.StartDateTime, destination.StartDateTime))
                    {
                        additionalData.Add("StartDateTime", origin.StartDateTime);
                    }
                    #endregion
                    if (!String.Equals(origin.Url, destination.Url))
                    {
                        additionalData.Add("Url", origin.Url);
                    }
                    if (!String.Equals(origin.Priority, destination.Priority))
                    {
                        additionalData.Add("Priority", origin.Priority);
                    }

                    if (additionalData.Keys.Count > 0)
                    {
                        FieldValueSet fieldsChange = new FieldValueSet();
                        fieldsChange.AdditionalData = additionalData;

                        try
                        {
                            var u = await GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                            upd++;
                            lblMessage.Text = "Task A: " + add + " D: " + del + " U:" + upd;
                        }
                        catch (Exception exUpd)
                        {
                            await DisplayMessageAsync($"Error Updating:{System.Environment.NewLine}{exUpd}");
                            App.logger.Error(exUpd.Message);
                        }                        
                    }
                }
            }
            #endregion

            App.logger.Information("Task Added: " + add + " Deleted: " + del + " Updated:" + upd);
        }


        private async Task ConciliationAssignments(Dictionary<string, MetaPlannerAssignment> sharePointAssignees, Dictionary<string, string> itemIds, Dictionary<string, ListItem> items)
        {
            int add = 0;
            int del = 0;

            #region Add from planner not in sharepoint
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerAssignment> entry in PlannerAssignments)
            {
                if (!sharePointAssignees.ContainsKey(entry.Key))
                {
                    var assigneeItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", entry.Value.TaskId},
                                {"UserId", entry.Value.UserId},
                            }
                        }
                    };
                    try
                    {
                        var a = await GraphClient.Sites[config.Site].Lists["assignees"].Items.Request().AddAsync(assigneeItem);
                        add++;
                        lblMessage.Text = "Assignees Added: " + add + " Deleted: " + del;
                    }
                    catch (Exception exAdd)
                    {
                        await DisplayMessageAsync($"Error Adding:{System.Environment.NewLine}{exAdd}");
                        App.logger.Error(exAdd.Message);
                    }
                }
            }
            #endregion

            #region Delete in SharePoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerAssignment> entry in sharePointAssignees)
            {
                if (!PlannerAssignments.ContainsKey(entry.Key))
                {
                    try
                    {
                        await GraphClient.Sites[config.Site].Lists["assignees"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        del++;
                        lblMessage.Text = "Assignees Added: " + add + " Deleted: " + del;
                    }
                    catch (Exception exDel)
                    {
                        await DisplayMessageAsync($"Error Deleting:{System.Environment.NewLine}{exDel}");
                        App.logger.Error(exDel.Message);
                    }
                }
            }
            #endregion

            lblMessage.Text = "Assignees Added: " + add + " Deleted: " + del ;
            App.logger.Information("Assignees Added: " + add + " Deleted: " + del);
        }
        #endregion


        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGroupButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO delete
            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);
            RadDataGrid.StartBringIntoView();

            try
            {
                // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                //var users = await graphClient.Users.Request().GetAsync();

                var plans = await GraphClient.Me.Planner.Plans.Request().GetAsync();

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
                    var group = await GraphClient.Groups[p.Owner].Request().GetAsync();

                    listPlan.Add(new MetaPlannerPlan()
                    {
                        PlanId = p.Id,
                        PlanName = p.Title,
                        CreatedBy = p.CreatedBy.User.Id,
                       // CreatedDate = (DateTime) p.CreatedDateTime,
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
                               // {"CreatedDate",  (DateTime)p.CreatedDateTime},
                                {"GroupName",  group.DisplayName},
                                {"GroupDescription",  group.Description},
                                {"GroupMail",  group.Mail},
                                {"Url", "https://tasks.office.com/congenrep.onmicrosoft.com/Home/PlanViews/"+p.Id}
                            }
                        }
                    };
                    await GraphClient.Sites[config.Site].Lists["plans"].Items.Request().AddAsync(planItem);

                    counter++;


                    var buckets = await GraphClient.Planner.Plans[p.Id].Buckets.Request().GetAsync();

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
                        await GraphClient.Sites[config.Site].Lists["buckets"].Items.Request().AddAsync(bucketItem);

                    }

                    var pTasks = await GraphClient.Planner.Plans[p.Id].Tasks.Request().GetAsync();


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
                        await GraphClient.Sites[config.Site].Lists["tasks"].Items.Request().AddAsync(taskItem);

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
                            await GraphClient.Sites[config.Site].Lists["assignees"].Items.Request().AddAsync(assigneesItem);

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
                await writer.Write(listPlan, storageFolder, prefix+" plans.csv");

                //writer.Write(listBuckets, storageFolder, "buckets.csv");
                await writer.Write(listBuckets, storageFolder, prefix + " buckets.csv");

                //writer.Write(listTasks, storageFolder, "tasks.csv");
               await writer.Write(listTasks, storageFolder, prefix + " tasks.csv");

                //writer.Write(listAssignment, storageFolder, "assignees.csv");
               await writer.Write(listAssignment, storageFolder, prefix + " assignees.csv");

                //File.Copy(sourceFile, destinationFile, true);
                 
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);

            }
            catch (MsalException msalEx)
            {
                await DisplayMessageAsync($"Error Acquiring Token:{System.Environment.NewLine}{msalEx}");
                App.logger.Error(msalEx.Message);
            }
            catch (Exception ex)
            {
                await DisplayMessageAsync($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
                App.logger.Error(ex.Message);
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
                App.logger.Error(ex.Message);
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
                App.logger.Error(ex.Message);
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
