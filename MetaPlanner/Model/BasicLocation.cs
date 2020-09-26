using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class BasicLocation
    {
        public BasicLocation()
        {
        }

        public BasicLocation(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out ubicacion);
        }

        private object ubicacion;
        public string Location
        {
            get
            {
                if (ubicacion != null)
                    return ubicacion.ToString();
                else
                    return null;
            }
            set { ubicacion = value; }
        }

        public string State
        {
            get
            {
                if (ubicacion != null)
                    return ubicacion.ToString().Split("-")[1].Trim();
                else
                    return null;
            }
        }

        public string City
        {
            get
            {
                if (ubicacion != null)
                    return ubicacion.ToString().Split("-")[0].Trim();
                else
                    return null;
            }
        }

    }
}
