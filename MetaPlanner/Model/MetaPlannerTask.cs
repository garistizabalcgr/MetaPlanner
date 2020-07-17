using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerTask
    {

        public MetaPlannerTask()
        {
        }

        public MetaPlannerTask(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out taskId);
            fields.TryGetValue("PlanId", out planId);
            fields.TryGetValue("BucketId", out bucketId);
            fields.TryGetValue("TaskName", out taskName);
            fields.TryGetValue("ActiveChecklistItemCount", out activeChecklistItemCount);
            fields.TryGetValue("AdditionalData", out additionalData);
            fields.TryGetValue("AssigneePriority", out assigneePriority);
            fields.TryGetValue("ConversationThreadId", out conversationThreadId);
            fields.TryGetValue("OrderHint", out orderHint);
            fields.TryGetValue("PercentComplete", out percentComplete);
            fields.TryGetValue("StartDateTime", out startDateTime);
            fields.TryGetValue("CreatedDateTime", out createdDateTime);
            fields.TryGetValue("CreatedBy", out createdBy);
            fields.TryGetValue("DueDateTime", out dueDateTime);
            fields.TryGetValue("HasDescription", out hasDescription);
            fields.TryGetValue("CompletedDateTime", out completedDateTime);
            fields.TryGetValue("ReferenceCount", out referenceCount);
            fields.TryGetValue("ChecklistItemCount", out checklistItemCount);
            fields.TryGetValue("Category1", out category1);
            fields.TryGetValue("Category2", out category2);
            fields.TryGetValue("Category3", out category3);
            fields.TryGetValue("Category4", out category4);
            fields.TryGetValue("Category5", out category5);
            fields.TryGetValue("Category6", out category6);
            fields.TryGetValue("Prefix", out prefix);
            fields.TryGetValue("Hours", out hours);
            fields.TryGetValue("CompletedBy", out completedBy);
            fields.TryGetValue("AssignmentsCount", out assignmentsCount);
            fields.TryGetValue("Url", out url);
        }



        private object taskId;
        public string TaskId
        {
            get {
                if (taskId != null)
                    return taskId.ToString();
                else
                    return null;
            }
            set { taskId = value; }
        }

        private object planId;
        public string PlanId
        {
            get {
                if (planId != null)
                    return planId.ToString();
                else
                    return null;
            }
            set { planId = value; }
        }

        private object bucketId;
        public string BucketId
        {
            get {
                if (bucketId != null)
                    return bucketId.ToString();
                else
                    return null;
            }
            set { bucketId = value; }
        }

        private object taskName;
        public string TaskName
        {
            get {
                if (taskName != null)
                    return taskName.ToString();
                else
                    return null;
            }
            set { taskName = value; }
        }

        private object activeChecklistItemCount;
        public string ActiveChecklistItemCount
        {
            get {
                if (activeChecklistItemCount != null)
                    return activeChecklistItemCount.ToString();
                else
                    return null;
            }
            set { activeChecklistItemCount = value; }
        }

        private object additionalData;
        public string AdditionalData
        {
            get {
                if (additionalData != null)
                    return additionalData.ToString();
                else
                    return null;
            }
            set { additionalData = value; }
        }

        private object assigneePriority;
        public string AssigneePriority
        {
            get {
                if (assigneePriority != null)
                    return assigneePriority.ToString();
                else
                    return null;
            }
            set { assigneePriority = value; }
        }

        private object conversationThreadId;
        public string ConversationThreadId
        {
            get {
                if (conversationThreadId != null)
                    return conversationThreadId.ToString();
                else
                    return null;
                }
            set { conversationThreadId = value; }
        }

        private object orderHint;
        public string OrderHint
        {
            get {
                if (orderHint != null)
                    return orderHint.ToString();
                else
                    return null;
            }
            set { orderHint = value; }
        }

    private object percentComplete;
    public string PercentComplete
    {
        get
            {
                if (percentComplete != null)
                    return percentComplete.ToString();
                else
                    return null;
            }
        set { percentComplete = value; }
    }

    private object startDateTime;
    public string StartDateTime
    {
        get
            {
                if (startDateTime != null)
                    return startDateTime.ToString();
                else
                    return null;
            }
        set { startDateTime = value; }
    }

    private object createdDateTime;
    public string CreatedDateTime
    {
        get
            {
                if (createdDateTime != null)
                    return createdDateTime.ToString();
                else
                    return null;
            }
        set { createdDateTime = value; }
    }

    private object createdBy;
    public string CreatedBy
    {
        get
            {
                if (createdBy != null)
                    return createdBy.ToString();
                else
                    return null;
            }
        set { createdBy = value; }
    }

        private object dueDateTime;
        public string DueDateTime
        {
            get
            {
                if (dueDateTime != null)
                    return dueDateTime.ToString();
                else
                    return null;
            }
            set { dueDateTime = value; }
        }

        private object hasDescription;
        public string HasDescription
        {
            get
            {
                if (hasDescription != null)
                    return hasDescription.ToString();
                else
                    return null;
            }
            set { hasDescription = value; }
        }

        private object completedDateTime;
        public string CompletedDateTime
        {
            get
            {
                if (completedDateTime != null)
                    return completedDateTime.ToString();
                else
                    return null;
            }
            set { completedDateTime = value; }
        }

        private object referenceCount;
        public string ReferenceCount
        {
            get
            {
                if (referenceCount != null)
                    return referenceCount.ToString();
                else
                    return null;
            }
            set { referenceCount = value; }
        }

        private object checklistItemCount;
        public string ChecklistItemCount
        {
            get
            {
                if (checklistItemCount != null)
                    return checklistItemCount.ToString();
                else
                    return null;
            }
            set { checklistItemCount = value; }
        }

        private object category1;
        public string Category1
        {
            get
            {
                if (category1 != null)
                    return category1.ToString();
                else
                    return null;
            }
            set { category1 = value; }
        }

        private object category2;
        public string Category2
        {
            get
            {
                if (category2 != null)
                    return category2.ToString();
                else
                    return null;
            }
            set { category2 = value; }
        }

        private object category3;
        public string Category3
        {
            get
            {
                if (category3 != null)
                    return category3.ToString();
                else
                    return null;
            }
            set { category3 = value; }
        }
        private object category4;
        public string Category4
        {
            get
            {
                if (category4 != null)
                    return category4.ToString();
                else
                    return null;
            }
            set { category4 = value; }
        }
        private object category5;
        public string Category5
        {
            get
            {
                if (category5 != null)
                    return category5.ToString();
                else
                    return null;
            }
            set { category5 = value; }
        }
        private object category6;
        public string Category6
        {
            get
            {
                if (category6 != null)
                    return category6.ToString();
                else
                    return null;
            }
            set { category6 = value; }
        }
        private object prefix;
        public string Prefix
        {
            get
            {
                if (prefix != null)
                    return prefix.ToString();
                else
                    return null;
            }
            set { prefix = value; }
        }
        private object hours;
        public string Hours
        {
            get
            {
                if (hours != null)
                    return hours.ToString();
                else
                    return null;
            }
            set { hours = value; }
        }
        private object completedBy;
        public string CompletedBy
        {
            get { 
                if (completedBy != null)
                    return completedBy.ToString();
                else
                    return null;
            }
            set { completedBy = value; }
        }
        private object assignmentsCount;
        public string AssignmentsCount
        {
            get
            {
                if (assignmentsCount != null)
                    return assignmentsCount.ToString();
                else
                    return null;
            }
            set { assignmentsCount = value; }
        }

        private object url;
        public string Url
        {
            get
            {
                if (url != null)
                    return url.ToString();
                else
                    return null;
            }
            set { url = value; }
        }
    }
}
