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
            fields.TryGetValue("Title", out userId);
            fields.TryGetValue("UserPrincipalName", out userPrincipalName);
            fields.TryGetValue("Mail", out mail);
            fields.TryGetValue("TheName", out theName);
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

        private object theName;
        public string TheName
        {
            get
            {
                if (theName != null)
                    return theName.ToString();
                else
                    return null;
            }
            set { theName = value; }
        }

    }
}
