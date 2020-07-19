using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerUser
    {

        public MetaPlannerUser()
        {
        }

        public MetaPlannerUser(IDictionary<string, object> fields)
        {
            fields.TryGetValue("UserId", out userId);
            fields.TryGetValue("UserPrincipalName", out userPrincipalName);
            fields.TryGetValue("Mail", out mail);
            fields.TryGetValue("DisplayName", out displayName);
            fields.TryGetValue("Department", out department);
            fields.TryGetValue("JobTitle", out jobTitle);
        }


        private object userId;
        public string UserId
        {
            get
            {
                if (userId != null)
                    return userId.ToString();
                else
                    return null;
            }
            set { userId = value; }
        }

        private object userPrincipalName;
        public string UserPrincipalName
        {
            get
            {
                if (userPrincipalName != null)
                    return userPrincipalName.ToString();
                else
                    return null;
            }
            set { userPrincipalName = value; }
        }

        private object mail;
        public string Mail
        {
            get
            {
                if (mail != null)
                    return mail.ToString();
                else
                    return null;
            }
            set { mail = value; }
        }

        private object displayName;
        public string DisplayName
        {
            get
            {
                if (displayName != null)
                    return displayName.ToString();
                else
                    return null;
            }
            set { displayName = value; }
        }


        private object department;
        public string Department
        {
            get
            {
                if (department != null)
                    return department.ToString();
                else
                    return null;
            }
            set { department = value; }
        }

        private object jobTitle;
        public string JobTitle
        {
            get
            {
                if (jobTitle != null)
                    return jobTitle.ToString();
                else
                    return null;
            }
            set { jobTitle = value; }
        }
    }
}
