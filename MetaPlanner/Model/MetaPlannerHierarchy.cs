using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{

    public class MetaPlannerHierarchy
    {
        public MetaPlannerHierarchy()
        {
        }

        public MetaPlannerHierarchy(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out planId);
            fields.TryGetValue("PlanName", out planName);
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
