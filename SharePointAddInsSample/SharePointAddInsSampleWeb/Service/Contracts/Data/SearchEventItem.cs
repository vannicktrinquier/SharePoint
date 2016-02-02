using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SharePointAddInsSampleWeb.Service.Contracts.Data
{
    public class SearchEventItem : SearchItem
    {
        /// <summary>
        /// Web url of the item
        /// </summary>
        public string WebUrl { get; set; }
        
        /// <summary>
        /// Web title of the item
        /// </summary>
        public string WebTitle { get; set; }
        
        /// <summary>
        /// Start date
        /// </summary>
        public DateTime? StartDate { get; set; }

        /// <summary>
        /// Due date
        /// </summary>
        public DateTime? EndDate { get; set; }

        /// <summary>
        /// Name of the person whom the task is assigned
        /// </summary>
        public string AssignedTo { get; set; }

        /// <summary>
        /// Percentage of the completion
        /// </summary>
        public string Location { get; set; }


    }
}
