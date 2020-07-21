using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerAssignment
    {
        public MetaPlannerAssignment()
        {
        }

        public MetaPlannerAssignment(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out taskId);
            fields.TryGetValue("UserId", out userId);
        }

        private object taskId;
        public string TaskId
        {
            get {
                if (taskId != null)
                    return taskId.ToString();
                else
                    return "";
            }
            set { taskId = value; }
        }

        private object userId;
        public string UserId
        {
            get
            {
                if (userId != null)
                    return userId.ToString();
                else
                    return "";
            }
            set { userId = value; }
        }

    }
}
