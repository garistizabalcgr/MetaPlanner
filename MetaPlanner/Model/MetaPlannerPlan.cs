using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    public class MetaPlannerPlan
    {

        public MetaPlannerPlan()
        {
        }

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


        private object planId;
        public string PlanId
        {
            get
            {
                if (planId != null)
                    return planId.ToString();
                else
                    return null;
            }
            set { planId = value; }
        }

        private object planName;
        public string PlanName
        {
            get
            {
                if (planName != null)
                    return planName.ToString();
                else
                    return null;
            }
            set { planName = value; }
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

        private object createdDate;
        public DateTimeOffset? CreatedDate
        {
            get
            {
                if (createdDate != null)
                    return DateTimeOffset.Parse(createdDate.ToString());
                else
                    return null;
            }
            set { createdDate = value; }
        }

        private object groupName;
        public string GroupName
        {
            get
            {
                if (groupName != null)
                    return groupName.ToString();
                else
                    return null;
            }
            set { groupName = value; }
        }

        private object groupDescription;
        public string GroupDescription
        {
            get
            {
                if (groupDescription != null)
                    return groupDescription.ToString();
                else
                    return null;
            }
            set { groupDescription = value; }
        }

        private object groupMail;
        public string GroupMail
        {
            get
            {
                if (groupMail != null)
                    return groupMail.ToString();
                else
                    return null;
            }
            set { groupMail = value; }
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
        public string Visible
        {
            get
            {
                if (visible != null)
                    return visible.ToString();
                else
                    return null;
            }
            set { visible = value; }
        }
    }
}
