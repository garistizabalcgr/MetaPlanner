using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    public class PlanManager
    {
        private static Random random = new Random();
        private static string[] planNames = new string[] { "Melbourne", "Sydney", "Brisbane", "Adelaide", "Perth" };
        public static List<Plan> plans = new List<Plan>();

        public static List<Plan> GetPlans()
        {
            return plans;
        }

        public static void SetPlans(List<Plan> list)
        {
            plans = list;
        }

    }
}
