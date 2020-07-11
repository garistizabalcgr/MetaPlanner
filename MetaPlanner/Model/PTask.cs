using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class PTask
    {
        public string TaskId { get; set; }
        public string TaskName { get; set; }
        public string PlanId { get; set; }
        public string BucketId { get; set; }
        public string ActiveChecklistItemCount { get; set; }
        public string AdditionalData { get; set; }
        public string AssigneePriority { get; set; }
        public string ConversationThreadId { get; set; }
        public string OrderHint { get; set; }
        //public string Details { get; set; }
        public string PercentComplete { get; set; }
        public string StartDateTime { get; set; }
        public string CreatedDateTime { get; set; }
        public string CreatedBy { get; set; }
        //public string ModifiedDateTime { get; set; }
        //public string ModifiedBy { get; set; }
        public string DueDateTime { get; set; }
        public string HasDescription { get; set; }
        public string CompletedDateTime { get; set; }
        public string ReferenceCount { get; set; }
        public string ChecklistItemCount { get; set; }
        
        public string Category1 { get; set; }
        public string Category2 { get; set; }
        public string Category3 { get; set; }
        public string Category4 { get; set; }
        public string Category5 { get; set; }
        public string Category6 { get; set; }
        public string Prefix { get; set; }
        public string Hours { get; set; }
        public string CompletedBy { get; set; }
        public string AssignmentsCount { get; set; }
        public string Url { get; set; }
        
    }
}
