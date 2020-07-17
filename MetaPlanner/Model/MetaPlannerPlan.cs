using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    public class MetaPlannerPlan
    {
        

        public MetaPlannerPlan(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out planId);
            fields.TryGetValue("PlanName", out planName);
            fields.TryGetValue("CreatedBy", out createdBy);
            fields.TryGetValue("CreatedDate", out createdDate);
            fields.TryGetValue("GroupName", out groupName);
            fields.TryGetValue("GroupDescription", out groupDescription);
            fields.TryGetValue("GroupMail", out groupMail);
            fields.TryGetValue("Url", out url);
            fields.TryGetValue("ParentId", out parentId);
            fields.TryGetValue("Visible", out visible);
        }

        public MetaPlannerPlan()
        {

        }

        private object planId;
        public string PlanId
        {
            get { return planId.ToString(); }  
            set { planId = value; }  
        }

        private object planName;
        public string PlanName
        {
            get { return planName.ToString(); }
            set { planName = value; }
        }

        private object createdBy;
        public string CreatedBy
        {
            get { return createdBy.ToString(); }
            set { createdBy = value; }
        }

        private object createdDate;
        public DateTimeOffset? CreatedDate
        {
            get{
                if (createdDate.GetType() == typeof(string) && createdDate != null)
                    return DateTimeOffset.Parse(createdDate.ToString());
                else if (createdDate != null)
                    return (DateTimeOffset)createdDate;
                else
                    return DateTimeOffset.MinValue; 
            }
            set {
                if (value.GetType() == typeof(string) )
                    createdDate = DateTimeOffset.Parse(value.ToString());
                else
                    createdDate = value; 
            }
        }

        private object groupName;
        public string GroupName
        {
            get { return groupName.ToString(); }
            set { groupName = value; }
        }

        private object groupDescription;
        public string GroupDescription
        {
            get { return groupDescription.ToString(); }
            set { groupDescription = value; }
        }

        private object groupMail;
        public string GroupMail
        {
            get { return groupMail.ToString(); }
            set { groupMail = value; }
        }

        private object url;
        public string Url
        {
            get { return url.ToString(); }
            set { url = value; }
        }

        private object parentId;
        public string ParentId
        {
            get
            {
                if (parentId != null)
                    return parentId.ToString();
                else

                    return null; 
            }
            set { parentId = value; }
        }

        private object visible;
        public bool Visible
        {
            get
            {
                if (parentId != null)
                    return Convert.ToBoolean(visible);
                else

                    return false;
            }
            set { parentId = value; }
        }

    }
}
