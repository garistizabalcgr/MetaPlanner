using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class BasicUser
    {

        public BasicUser()
        {
        }

        private object fullName;
        public string NombreCompleto
        {
            get
            {
                if (fullName != null)
                    return fullName.ToString();
                else
                    return null;
            }
            set { fullName = value; }
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

        private object upn;
        public string UserPrincipalName
        {
            get
            {
                if (upn != null)
                    return upn.ToString();
                else
                    return null;
            }
            set { upn = value; }
        }

    }
}
