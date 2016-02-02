using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SharePointAddInsSampleWeb.Service.Contracts.Data
{
    public class SearchWebItem : SearchItem
    {
        /// <summary>
        /// Last modified date
        /// </summary>
        public DateTime LastModifiedDate { get; set; }

        /// <summary>
        /// Rank of the item
        /// </summary>
        public float Rank{ get; set; }

        /// <summary>
        /// Status of the item
        /// </summary>
        public string Status{ get; set; }

    }
}
