using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class BasicUser
    {       public BasicUser()
        {
        }

        public BasicUser(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out upn);
            fields.TryGetValue("NombreCompleto", out fullName);
            fields.TryGetValue("Cargo", out jobTitle);
            fields.TryGetValue("Dependencia", out department);
            fields.TryGetValue("Sede", out location);
        }

        private object fullName;
        public string NombreCompleto
        {
            get
            {
                if (fullName != null)
                    return fullName.ToString();
                else
                    return "";
            }
            set { fullName = value; }
        }


        private object upn;
        public string UserPrincipalName
        {
            get
            {
                if (upn != null)
                    return upn.ToString();
                else
                    return "";
            }
            set { upn = value; }
        }

        private object jobTitle;
        public string Cargo
        {
            get
            {
                if (jobTitle != null)
                    return jobTitle.ToString();
                else
                    return "";
            }
            set { jobTitle = value; }
        }

        private object department;
        public string Dependencia
        {
            get
            {
                if (department != null)
                    return department.ToString();
                else
                    return "";
            }
            set { department = value; }
        }

        private object location;
        public string Sede
        {
            get
            {
                if (location != null)
                    return location.ToString();
                else
                    return "";
            }
            set { location = value; }
        }
    }
}
