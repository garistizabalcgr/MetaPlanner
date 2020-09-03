using MetaPlanner.Model;
using MetaPlanner.Output;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;
using System.Net.Http.Headers;

namespace MetaPlanner.Control
{
    class Command
    {
        //Main View
        private MainPage mainPage;

        //Dictionary of Plans
        public Dictionary<string, MetaPlannerPlan> PlannerPlans = new Dictionary<string, MetaPlannerPlan>();

        //Dictionary of Buckets
        public Dictionary<string, MetaPlannerBucket> PlannerBuckets = new Dictionary<string, MetaPlannerBucket>();

        //Dictionary of Tasks
        public Dictionary<string, MetaPlannerTask> PlannerTasks = new Dictionary<string, MetaPlannerTask>();

        //Dictionary of Assignments
        public Dictionary<string, MetaPlannerAssignment> PlannerAssignments = new Dictionary<string, MetaPlannerAssignment>();

        //Dictionary of Users
        public Dictionary<string, MetaPlannerUser> PlannerUsers = new Dictionary<string, MetaPlannerUser>();

        //Client to Cloud Service
        private GraphServiceClient GraphClient;

        // Configuration
        public static Configuration config = Configuration.ReadFromJsonFile();

        // The MSAL Public client GraphServiceClient
        private static IPublicClientApplication PublicClientApp;

        // private static string MSGraphURL = "https://graph.microsoft.com/v1.0/";
        private static AuthenticationResult authResult;

        //Folder of file storage
        private StorageFolder storageFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
        //Date of creation
        private String TimeStamp
        {
            get
            {
                return String.Format("{0:D4}", DateTime.Now.Year) + "-" + String.Format("{0:D2}", DateTime.Now.Month) + "-" + String.Format("{0:D2}", DateTime.Now.Day) + "_" + String.Format("{0:D2}", DateTime.Now.Hour) + "_" + String.Format("{0:D2}", DateTime.Now.Minute) + "_" + String.Format("{0:D2}", DateTime.Now.Second);
            }
        }
        //Cvs writer of files
        private Writer writer = new Writer();

        public Command(MainPage mainPage)
        {
            this.mainPage = mainPage;
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

            mainPage.DisplayMessage("GetPlannerPlans getting " + listPlanner.Count);

            //Get hierarchy from Sharepoint
            var listHierarchy = await GetSharePointList("hierarchy"); 
            Dictionary<string, MetaPlannerHierarchy> sharePointHierarchy = new Dictionary<string, MetaPlannerHierarchy>();
            foreach (ListItem item in listHierarchy)
            {
                MetaPlannerHierarchy hierarchy = new MetaPlannerHierarchy(item.Fields.AdditionalData);
                sharePointHierarchy.Add(hierarchy.PlanId, hierarchy);
            }

            PlannerPlans = new Dictionary<string, MetaPlannerPlan>();
            foreach (PlannerPlan p in listPlanner)
            {
                MetaPlannerHierarchy theHierarchy;
                if (!sharePointHierarchy.TryGetValue(p.Id, out theHierarchy))
                {
                    theHierarchy = new MetaPlannerHierarchy();
                    theHierarchy.PlanId = p.Id;
                    theHierarchy.ParentId = null;
                    theHierarchy.Visible = "true";
                }

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
                        Url = "https://tasks.office.com/" + config.Tenant + "/Home/PlanViews/" + p.Id,
                        ParentId = theHierarchy.ParentId,
                        Visible = theHierarchy.Visible
                    });
            }

            App.logger.Information("GetPlannerPlans Total: " + PlannerPlans.Count);
        }

        /// <summary>
        /// Process Plans Data
        /// </summary>
        public async Task ProcessPlans()
        {
                 // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                await GetPlannerPlans();

                await WriteAndUpload(PlannerPlans, "plans");

                if (config.IsSharePointListEnabled)
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
                if (!sharePointPlans.ContainsKey(entry.Key))
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
                                {"Url", entry.Value.Url},
                                {"ParentId",entry.Value.ParentId },
                                {"Visible",entry.Value.Visible }
                            }
                        }
                    };
                    if (add % config.ChunkSize != 0)
                    {
                        var a = GraphClient.Sites[config.Site].Lists["plans"].Items.Request().AddAsync(planItem);
                    }
                    else
                    {
                        var a = await GraphClient.Sites[config.Site].Lists["plans"].Items.Request().AddAsync(planItem);
                    }
                    add++;
                    mainPage.DisplayMessage( "Plan A: " + add + " D: " + del + " U:" + upd);
                }
            }
            #endregion

            #region Delete in sharepoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerPlan> entry in sharePointPlans)
            {
                if (!PlannerPlans.ContainsKey(entry.Key))
                {
                    if (del % config.ChunkSize != 0)
                    {
                        GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    }
                    else
                    {
                        await GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    }
                    del++;
                    mainPage.DisplayMessage("Plan A: " + add + " D: " + del + " U:" + upd);
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
                    if (!String.Equals(origin.GroupDescription, destination.GroupDescription))
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
                    if (!String.Equals(origin.ParentId, destination.ParentId))
                    {
                        additionalData.Add("ParentId", origin.ParentId);
                    }
                    if (!String.Equals(origin.Visible, destination.Visible))
                    {
                        additionalData.Add("Visible", origin.Visible);
                    }

                    //CreateDate is readOnly
                    if (additionalData.Keys.Count > 0)
                    {
                        FieldValueSet fieldsChange = new FieldValueSet();
                        fieldsChange.AdditionalData = additionalData;

                        if (upd % config.ChunkSize != 0)
                        {
                            object obj = GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        }
                        else
                        {
                           object obj = await GraphClient.Sites[config.Site].Lists["plans"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        }
                        upd++;
                        mainPage.DisplayMessage("Plan A: " + add + " D: " + del + " U:" + upd);
                    }
                }
            }
            #endregion
            
            App.logger.Information("ConciliationPlans Plan Added: " + add + " Deleted: " + del + " Updated:" + upd);
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

                mainPage.DisplayMessage("GetPlannerBuckets getting " + listBuckets.Count);

                foreach (PlannerBucket bucket in listBuckets)
                {
                    PlannerBuckets.Add(bucket.Id, new MetaPlannerBucket()
                    {
                        BucketId = bucket.Id,
                        BucketName = bucket.Name,
                        OrderHint = bucket.OrderHint,
                        PlanId = plan.PlanId
                    });
                }
            }
            App.logger.Information("GetPlannerBuckets Total: " + PlannerBuckets.Count);
        }

        /// <summary
        /// Process Buckets Data
        /// </summary>
        public async Task ProcessBuckets()
        {
            // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                await GetPlannerBuckets();
                await WriteAndUpload(PlannerBuckets, "buckets");

                if (config.IsSharePointListEnabled)
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
                    if (add % config.ChunkSize != 0)
                    {
                        object obj = GraphClient.Sites[config.Site].Lists["buckets"].Items.Request().AddAsync(bucketItem);
                    }
                    else
                    {
                        object obj = await GraphClient.Sites[config.Site].Lists["buckets"].Items.Request().AddAsync(bucketItem);
                    }
             
                    add++;
                    mainPage.DisplayMessage("Bucket A: " + add + " D: " + del + " U:" + upd);
                }
            }
            #endregion

            #region Delete in SharePoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerBucket> entry in sharePointBuckets)
            {
                if (!PlannerBuckets.ContainsKey(entry.Key))
                {
                    await GraphClient.Sites[config.Site].Lists["buckets"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    del++;
                    mainPage.DisplayMessage("Bucket A: " + add + " D: " + del + " U:" + upd);
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

                        if (upd % config.ChunkSize != 0)
                        {
                            object obj = GraphClient.Sites[config.Site].Lists["buckets"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        }
                        else
                        {
                            object obj = await GraphClient.Sites[config.Site].Lists["buckets"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        }
                            
                        upd++;
                        mainPage.DisplayMessage("Bucket A: " + add + " D: " + del + " U:" + upd);
                    }
                }
            }
            #endregion

            App.logger.Information("ConciliationBuckets Buckets Added: " + add + " Deleted: " + del + " Updated:" + upd);
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
                mainPage.DisplayMessage("GetPlannerTasks getting " + listTasks.Count);

                foreach (PlannerTask task in listTasks)
                {
                    MetaPlannerTask myTask = new MetaPlannerTask() { TaskId = task.Id };

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
                            try { 
                            myTask.Hours = float.Parse(two.Substring(0, k).Trim().Replace('.', ','));
                            myTask.TaskName = two.Substring(k + 1).Trim();
                            }
                            catch(Exception badNumber)
                            {
                                myTask.Hours = -1;
                                myTask.TaskName = two.Trim();
                            }
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
                    myTask.CompletedDateTime = task.CompletedDateTime;
                    myTask.ConversationThreadId = task.ConversationThreadId;
                    myTask.CreatedBy = task.CreatedBy.User.Id;
                    myTask.CreatedDateTime = task.CreatedDateTime;
                    myTask.DueDateTime = task.DueDateTime;
                    myTask.HasDescription = task.HasDescription.ToString();
                    myTask.OrderHint = task.OrderHint;
                    myTask.PercentComplete = task.PercentComplete.ToString();
                    myTask.ReferenceCount = task.ReferenceCount.ToString();
                    myTask.StartDateTime = task.StartDateTime;
                    myTask.Url = "https://tasks.office.com/" + config.Tenant + "/es-es/Home/Task/" + task.Id;
                    #endregion
                    object priority;
                    task.AdditionalData.TryGetValue("priority", out priority);
                    if (priority != null)
                        myTask.Priority = priority.ToString();
                    PlannerTasks.Add(myTask.TaskId, myTask);
                    GetPlannerAssignment(task);
                }
            }
            App.logger.Information("GetPlannerTasks Total: " + PlannerTasks.Count);
            App.logger.Information("GetPlannerAssignment Total: " + PlannerAssignments.Count);
        }

        /// <summary
        /// Process Task Data
        /// </summary>
        /// 
        private void GetPlannerAssignment(PlannerTask task)
        {
            foreach (string userId in task.Assignments.Assignees)
            {
                PlannerAssignments.Add(task.Id + "_" + userId, new MetaPlannerAssignment()
                {
                    TaskId = task.Id,
                    UserId = userId
                });
            }
            mainPage.DisplayMessage("GetPlannerAssignment getting " + task.Assignments.Count);
        }

        /// <summary
        /// Process Task Data
        /// </summary>
        public async Task ProcessTasks()
        {
                 // Sign-in user using MSAL and obtain an access token for MS Graph
                GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

                await GetPlannerTasks();

                await WriteAndUpload(PlannerTasks, "tasks");
                await WriteAndUpload(PlannerAssignments, "assignees");

                if (config.IsSharePointListEnabled)
                {

                    #region Get bulk data from SharePoint Tasks
                    var listTasks = await GetSharePointList("tasks");

                    Dictionary<string, MetaPlannerTask> sharePointTasks = new Dictionary<string, MetaPlannerTask>();
                    Dictionary<string, string> itemIds = new Dictionary<string, string>();
                    Dictionary<string, ListItem> items = new Dictionary<string, ListItem>();

                    foreach (ListItem item in listTasks)
                    {
                        MetaPlannerTask task = new MetaPlannerTask(item.Fields.AdditionalData);
                        if (task != null && task.TaskId != null)
                        {
                            sharePointTasks.Add(task.TaskId, task);
                        }
                        else
                        {
                            mainPage.DisplayMessage("Task is null");
                            App.logger.Error("Task is null");
                        }
                        if (item.Id != null && task.TaskId != null)
                        {
                            itemIds.Add(task.TaskId, item.Id);
                        }
                        else
                        {
                            mainPage.DisplayMessage("Task is null");
                            App.logger.Error("Task is null");
                        }
                        if (item != null && item.Id != null)
                        {
                            items.Add(item.Id, item);
                        }
                        else
                        {
                            mainPage.DisplayMessage("Task is null");
                            App.logger.Error("Task is null");
                        }
                }
                    #endregion
                     await ConciliationTasks(sharePointTasks, itemIds, items);

                    mainPage.DisplayMessage("ConciliationTasks called " + sharePointTasks.Count);
                #region Get bulk data from SharePoint
                var listAssignees = await GetSharePointList("assignees");

                    Dictionary<string, MetaPlannerAssignment> sharePointAsignees = new Dictionary<string, MetaPlannerAssignment>();
                    Dictionary<string, string> itemIdsA = new Dictionary<string, string>();
                    Dictionary<string, ListItem> itemsA = new Dictionary<string, ListItem>();

                    foreach (ListItem theItem in listAssignees)
                    {
                        MetaPlannerAssignment assignment = new MetaPlannerAssignment(theItem.Fields.AdditionalData);
                        sharePointAsignees.Add(assignment.TaskId + "_" + assignment.UserId, assignment);
                        itemIdsA.Add(assignment.TaskId + "_" + assignment.UserId, theItem.Id);
                        itemsA.Add(theItem.Id, theItem);
                    }
                    #endregion
                    await ConciliationAssignments(sharePointAsignees, itemIdsA, itemsA);
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
                        if (add % config.ChunkSize != 0)
                        {
                            var a =  GraphClient.Sites[config.Site].Lists["tasks"].Items.Request().AddAsync(taskItem);
                        }
                        else
                        {
                            var b =  await GraphClient.Sites[config.Site].Lists["tasks"].Items.Request().AddAsync(taskItem);
                        }
                        add++;
                        mainPage.DisplayMessage("Task A: " + add + " D: " + del + " U:" + upd);
                    }
                    catch (Exception exception)
                    {
                        mainPage.DisplayMessage($"Task Error Adding:{System.Environment.NewLine}{exception}" + taskItem.Name);
                        App.logger.Error("ConciliationTasks Add " + exception.Message);
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
                        if (del % config.ChunkSize != 0)
                        {
                            GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        }
                        else
                        {
                            await GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        }
                        del++;
                        mainPage.DisplayMessage("Task A: " + add + " D: " + del + " U:" + upd);
                    }
                    catch (Exception exDel)
                    {
                        mainPage.DisplayMessage($"Error Deleting:{System.Environment.NewLine}{exDel}");
                        App.logger.Error("ConciliationTasks Delete " + exDel.Message);
                    }
                }
            }
            #endregion


            #region Update in Sharepoint changes from planner
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerTask> mytask in PlannerTasks)
            {
                if (sharePointTasks.ContainsKey(mytask.Key))
                {
                    try
                    {
                        MetaPlannerTask origin = PlannerTasks[mytask.Key];
                        MetaPlannerTask destination = sharePointTasks[mytask.Key];

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


                        if (!String.Equals(origin.ConversationThreadId, destination.ConversationThreadId))
                        {
                            additionalData.Add("ConversationThreadId", origin.ConversationThreadId);
                        }

                        if (!String.Equals(origin.CreatedBy, destination.CreatedBy))
                        {
                            additionalData.Add("CreatedBy", origin.CreatedBy);
                        }

                        if (DateTimeOffset.Compare((DateTimeOffset)origin.CreatedDateTime, (DateTimeOffset)destination.CreatedDateTime) != 0)
                        {
                            additionalData.Add("CreatedDateTime", origin.CreatedDateTime);
                        }

                        if (origin.StartDateTime != null && destination.StartDateTime != null && 
                            DateTimeOffset.Compare((DateTimeOffset)origin.StartDateTime, (DateTimeOffset)destination.StartDateTime) != 0)
                        {
                            additionalData.Add("StartDateTime", origin.StartDateTime);
                        }

                        if (origin.CompletedDateTime != null &&  destination.CompletedDateTime != null &&
                            DateTimeOffset.Compare((DateTimeOffset)origin.CompletedDateTime, (DateTimeOffset)destination.CompletedDateTime) != 0)
                        {
                            additionalData.Add("CompletedDateTime", origin.CompletedDateTime);
                        }

                        if (origin.DueDateTime != null &&  destination.DueDateTime != null &&
                            DateTimeOffset.Compare((DateTimeOffset)origin.DueDateTime, (DateTimeOffset)destination.DueDateTime) != 0)
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

                        if (!String.Equals(origin.Url, destination.Url))
                        {
                            additionalData.Add("Url", origin.Url);
                        }
                        #endregion

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
                                if (upd % config.ChunkSize != 0)
                                {
                                    var u =  GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[mytask.Key]].Fields.Request().UpdateAsync(fieldsChange);
                                }
                                else
                                {
                                    var u = await GraphClient.Sites[config.Site].Lists["tasks"].Items[itemIds[mytask.Key]].Fields.Request().UpdateAsync(fieldsChange);
                                }
                                upd++;
                                mainPage.DisplayMessage("Task A: " + add + " D: " + del + " U:" + upd);
                            }
                            catch (Exception exception)
                            {
                                mainPage.DisplayMessage($"Task Error Updating:{System.Environment.NewLine}{exception}" + mytask.Key);
                                App.logger.Error("ConciliationTasks Update " + exception.Message);
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        mainPage.DisplayMessage($"Task Error Updating:{System.Environment.NewLine}{exception}" + mytask.Key);
                        App.logger.Error("ConciliationTasks Update " + exception.Message);
                    }
                }
            }
            #endregion            

            mainPage.DisplayMessage("Task Added: " + add + " Deleted: " + del + " Updated: " + upd);
            App.logger.Information("ConciliationTasks Task Added: " + add + " Deleted: " + del + " Updated:" + upd);
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
                        if (add % config.ChunkSize != 0)
                        {

                            var a = GraphClient.Sites[config.Site].Lists["assignees"].Items.Request().AddAsync(assigneeItem);
                        }
                        else
                        {
                            var a = await GraphClient.Sites[config.Site].Lists["assignees"].Items.Request().AddAsync(assigneeItem);
                        }
                        add++;
                        mainPage.DisplayMessage("Assignees Added: " + add + " Deleted: " + del);
                    }
                    catch (Exception exception)
                    {
                        mainPage.DisplayMessage($"Error Adding:{System.Environment.NewLine}{exception}");
                        App.logger.Error("ConciliationAssignments Add " + exception.Message);
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
                        if (add % config.ChunkSize != 0)
                        {
                            GraphClient.Sites[config.Site].Lists["assignees"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        }
                        else
                        {
                           await GraphClient.Sites[config.Site].Lists["assignees"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                        }
                        del++;
                        mainPage.DisplayMessage("Assignees Added: " + add + " Deleted: " + del);
                    }
                    catch (Exception exception)
                    {
                        mainPage.DisplayMessage($"Error Deleting:{System.Environment.NewLine}{exception}");
                        App.logger.Error("ConciliationAssignments Delete " + exception.Message);
                    }
                }
            }
            #endregion

            mainPage.DisplayMessage("Assignees Added: " + add + " Deleted: " + del);
            App.logger.Information("ConciliationAssignments Assignees Added: " + add + " Deleted: " + del);
        }
        #endregion

        #region Users
        /// <summary>
        /// Get all data from Plans from Planner in plannerPlans
        /// </summary>
        private async Task GetPlannerUsers()
        {
            if (PlannerTasks == null || PlannerTasks.Count == 0)
            {
                await GetPlannerTasks();
            }


             //Get historic users from Sharepoint
            var listUsers = await GetSharePointList("users");
            Dictionary<string, MetaPlannerUser> sharePointUsers = new Dictionary<string, MetaPlannerUser>();
            foreach (ListItem item in listUsers)
            {
                MetaPlannerUser user = new MetaPlannerUser(item.Fields.AdditionalData);
                sharePointUsers.Add(user.UserId, user);
            }

            mainPage.DisplayMessage("GetPlannerUsers getting " + listUsers.Count);

            PlannerUsers = new Dictionary<string, MetaPlannerUser>();

            foreach (MetaPlannerAssignment assignee in PlannerAssignments.Values)
            {
                if (!PlannerUsers.ContainsKey(assignee.UserId))
                {
                    try
                    {
                        var user = await GraphClient.Users[assignee.UserId].Request().GetAsync();

                        PlannerUsers.Add(assignee.UserId, new MetaPlannerUser()
                        {
                            UserId = assignee.UserId,
                            UserPrincipalName = user.UserPrincipalName,
                            TheName = user.DisplayName,
                            Mail = user.Mail
                        });
                    }
                    catch (Exception exeption)
                    {
                        
                        //Get from historic data
                        MetaPlannerUser theUser;
                        if (sharePointUsers.TryGetValue(assignee.UserId, out theUser))
                        {
                            PlannerUsers.Add(theUser.UserId, theUser);
                        }
                        else 
                        //is totally forgotten
                        {
                            App.logger.Error("GetPlannerUsers " + exeption.Message);
                            PlannerUsers.Add(assignee.UserId, new MetaPlannerUser()
                            {
                                UserId = assignee.UserId,
                                UserPrincipalName = assignee.UserId,
                                TheName = assignee.UserId,
                                Mail = assignee.UserId
                            });
                        }
                    }
                }
            }

            mainPage.DisplayMessage("GetPlannerUsers Total: " + PlannerUsers.Count);
            App.logger.Information("GetPlannerUsers Total: " + PlannerUsers.Count);
        }

       
        
        public async Task ProcessUsers() {

            // Sign-in user using MSAL and obtain an access token for MS Graph
            GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);
            await GetPlannerUsers();
            await WriteAndUpload(PlannerUsers, "users");

            if (config.IsSharePointListEnabled)
            {
                #region Get bulk data from SharePoint
                var listUsers = await GetSharePointList("users");

                Dictionary<string, MetaPlannerUser> sharePointUsers = new Dictionary<string, MetaPlannerUser>();
                Dictionary<string, string> itemIds = new Dictionary<string, string>();
                Dictionary<string, ListItem> items = new Dictionary<string, ListItem>();
                foreach (ListItem item in listUsers)
                {
                    MetaPlannerUser user = new MetaPlannerUser(item.Fields.AdditionalData);
                    sharePointUsers.Add(user.UserId, user);
                    itemIds.Add(user.UserId, item.Id);
                    items.Add(item.Id, item);

                }
                #endregion
                await ConciliationUsers(sharePointUsers, itemIds, items);
            }

        }

        private async Task ConciliationUsers(Dictionary<string, MetaPlannerUser> sharePointUsers, Dictionary<string, string> itemIds, Dictionary<string, ListItem> items)
        {
            int add = 0;
            int del = 0;
            int upd = 0;

            #region Add from planner not in sharepoint
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerUser> entry in PlannerUsers)
            {
                if (!sharePointUsers.ContainsKey(entry.Key))
                {
                    var userItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"Title", entry.Value.UserId},
                                {"UserPrincipalName", entry.Value.UserPrincipalName},
                                {"Mail", entry.Value.Mail},
                                {"TheName", entry.Value.TheName }
                            }
                        }
                    };


                    if (add % config.ChunkSize != 0)
                    {
                        GraphClient.Sites[config.Site].Lists["users"].Items.Request().AddAsync(userItem);
                    }
                    else
                    {
                        await GraphClient.Sites[config.Site].Lists["users"].Items.Request().AddAsync(userItem);
                    }
                    add++;
                    mainPage.DisplayMessage("Users A: " + add + " D: " + del + " U:" + upd);
                }
            }
            #endregion

            /*
            #region Delete in sharepoint not in planner
            //Delete from SharePoint not in Planner
            foreach (KeyValuePair<string, MetaPlannerUser> entry in sharePointUsers)
            {
                if (!PlannerUsers.ContainsKey(entry.Key))
                {
                    GraphClient.Sites[config.Site].Lists["users"].Items[itemIds[entry.Key]].Request().DeleteAsync();
                    del++;
                    mainPage.DisplayMessage("Users A: " + add + " D: " + del + " U:" + upd);
                }
            }
            #endregion
            */
            /*
            #region Update in Sharepoint changes from planner
            //Add new from Planner to SharePoint
            foreach (KeyValuePair<string, MetaPlannerPlan> entry in PlannerPlans)
            {
                if (sharePointUsers.ContainsKey(entry.Key))
                {
                    MetaPlannerUser origin = PlannerUsers[entry.Key];
                    MetaPlannerUser destination = sharePointUsers[entry.Key];

                    Dictionary<string, object> additionalData = new Dictionary<string, object>();

                    if (!String.Equals(origin.UserPrincipalName, destination.UserPrincipalName))
                    {
                        additionalData.Add("UserPrincipalName", origin.UserPrincipalName);
                    }
                    if (!String.Equals(origin.Mail, destination.Mail))
                    {
                        additionalData.Add("Mail", origin.Mail);
                    }
                    if (!String.Equals(origin.TheName, destination.TheName))
                    {
                        additionalData.Add("TheName", origin.TheName);
                    }

                    if (additionalData.Keys.Count > 0)
                    {
                        FieldValueSet fieldsChange = new FieldValueSet();
                        fieldsChange.AdditionalData = additionalData;
                        object obj = GraphClient.Sites[config.Site].Lists["users"].Items[itemIds[entry.Key]].Fields.Request().UpdateAsync(fieldsChange);
                        upd++;
                        mainPage.DisplayMessage("Users A: " + add + " D: " + del + " U:" + upd);
                    }
                }
            }
            #endregion
            */
            App.logger.Information("ConciliationUsers Users Added: " + add + " Deleted: " + del + " Updated:" + upd);
        }
        #endregion


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
                App.logger.Error("SignInUserAndGetTokenUsingMSAL " + ex.Message);
                authResult = await PublicClientApp.AcquireTokenInteractive(scopes)
                                                  .ExecuteAsync()
                                                  .ConfigureAwait(false);

            }
            return authResult.AccessToken;
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

            int counter = 0; 
            foreach (ListItem item in allItems)
            {
                try
                {
                    counter++;
                    if (counter % config.ChunkSize != 0)
                    {
                         GraphClient.Sites[config.Site].Lists[listName].Items[item.Id].Request().DeleteAsync();
                    }
                    else
                    {
                        await GraphClient.Sites[config.Site].Lists[listName].Items[item.Id].Request().DeleteAsync();
                    }
                }
                catch (Exception exception)
                {
                    App.logger.Error("CleanSharepointList " + exception.Message);
                }
            }
            mainPage.DisplayMessage("CleanSharepointList " + listName + " " + allItems.Count);
            App.logger.Information("CleanSharepointList " + listName + " " + allItems.Count);
        }

        public async Task CleanAllSharePointLists()
        {
            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

            GraphClient = await SignInAndInitializeGraphServiceClient(config.ScopesArray);

            
            CleanSharepointList("tasks");
            CleanSharepointList("buckets");
            CleanSharepointList("plans");
            await CleanSharepointList("assignees");


            Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
        }

        private async Task WriteAndUpload(System.Collections.IDictionary dictionary, string name)
        {


            var drive = await GraphClient.Sites[config.Site].Drive.Request().GetAsync();
            
            //Check the Folder in SharePoint
            DriveItem root = await GraphClient.Sites[config.Site].Drives[drive.Id].Root.Request().GetAsync();
            IDriveItemChildrenCollectionPage children = await GraphClient.Drives[drive.Id].Items[root.Id].Children.Request().GetAsync();//TODO user root instead items
            DriveItem theFolder = children.Where(c => c.Name == config.FolderName).FirstOrDefault();
           
            
            // If does not exist create it
            if (theFolder == null)
            {
                theFolder = new DriveItem { 
                    Name = config.FolderName, 
                    Folder = new Folder()
                };
                theFolder =  await GraphClient.Drives[drive.Id].Root.Children.Request().AddAsync(theFolder);
            }


            //Prepare the stream to upload
            string fileName =  name + ".csv";
            await writer.Write(dictionary.Values, storageFolder, fileName);
            App.logger.Information("Write " + fileName + " " + dictionary.Values.Count + " lines");
            FileStream fileStream = new FileStream(storageFolder.Path + "\\" + fileName, FileMode.Open, FileAccess.Read);

            DriveItem theFile = new DriveItem() { 
                Name = fileName, 
                File = new Microsoft.Graph.File(),
            };

            try
            {
                //small file  < 4 Mb
                
                var uploadFile = await GraphClient.Sites[config.Site].
                    Drives[drive.Id].
                    Root.ItemWithPath(config.FolderName + "/" + fileName).
                    Content.Request().
                    PutAsync<DriveItem>(fileStream);

                mainPage.DisplayMessage("Ok "+ config.FolderName + "/" + fileName);

                //large file 
                /*
                var uploadSession = GraphClient.Sites[config.Site].
                    Drives[drive.Id].
                    Root.ItemWithPath(config.FolderName + "/" + fileName)
                    .CreateUploadSession()
                    .Request().PostAsync().Result;

                var maxChunkSize = 50 * 1024;
                var largeUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxChunkSize);
                IProgress<long> uploadProgress = new Progress<long>(uploadBytes =>
                {
                    mainPage.DisplayMessage($"Uploaded{uploadBytes} of {fileStream.Length} bytes");
                });
                UploadResult<DriveItem> uploadResult = largeUploadTask.UploadAsync(uploadProgress).Result;
                if (uploadResult.UploadSucceeded)
                {
                    mainPage.DisplayMessage("Uploaded Ok");
                }
                */
            }
            catch (Exception exception)
            {
                mainPage.DisplayMessage("Error WriteAndUpload " +exception.Message);
                App.logger.Error("WriteAndUpload " + exception.Message);
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
                    var uploadFile = await GraphClient.Sites[config.Site].Drives[drive.Id].Root.ItemWithPath(config.FolderName + "/" +config.SubFolderName+"/"+ fileNameT).Content.Request().PutAsync<DriveItem>(fsT);
                    mainPage.DisplayMessage("Ok " + config.FolderName + "/" + config.SubFolderName + "/" + fileNameT);
                }
                catch (Exception exception)
                {
                    mainPage.DisplayMessage(exception.Message);
                    App.logger.Error("WriteAndUpload " + exception.Message);
                }
            }
        }

    }
}
