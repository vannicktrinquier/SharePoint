using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SharePointAddInsSampleWeb.Service.Contracts.Data
{
    public class SearchDocumentItem : SearchItem
    {
        /// <summary>
        /// Link of the item
        /// </summary>
        public string OnlineLink { get; set; }

        /// <summary>
        /// Last modified date
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Last modified date
        /// </summary>
        public DateTime LastModifiedDate { get; set; }

        /// <summary>
        /// Name of the latest person that upade the item
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Name of the latest person that upade the item
        /// </summary>
        public string ModifiedBy { get; set; }


        /// <summary>
        /// Web url of the item
        /// </summary>
        public string WebUrl { get; set; }

        /// <summary>
        /// Rank of the item
        /// </summary>
        public float Rank{ get; set; }

    }
}
