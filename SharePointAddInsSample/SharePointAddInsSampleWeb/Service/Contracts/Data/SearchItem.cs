using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SharePointAddInsSampleWeb.Service.Contracts.Data
{
    public abstract class SearchItem
    {
        /// <summary>
        /// Title of the item
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Link of the item
        /// </summary>
        public string Link { get; set; }
    }
}
