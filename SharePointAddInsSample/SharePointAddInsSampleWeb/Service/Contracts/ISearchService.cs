using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SharePointAddInsSampleWeb.Service.Contracts.Data;

namespace SharePointAddInsSampleWeb.Service.Contracts
{
    public interface ISearchService
    {

        /// <summary>
        /// Retrieve the latest documents modified from the current site and all subsites
        /// </summary>
        /// <returns></returns>
        IEnumerable<SearchDocumentItem> GetLatestModifiedDocuments();

        /// <summary>
        /// Retrieve the most popular documents from the current site and all subsites
        /// </summary>
        /// <returns></returns>
        IEnumerable<SearchDocumentItem> GetMostPopularDocuments();

        /// <summary>
        /// Retrieve the most popular projects from the current site and all subsites
        /// </summary>
        /// <returns></returns>
        IEnumerable<SearchWebItem> GetMostPopularProjects();

        /// <summary>
        /// Retrieve the tasks assigned to me that have not been yet closed from the current site and all subsites
        /// </summary>
        /// <returns></returns>
        IEnumerable<SearchTaskItem> GetMyActiveTasks();

        /// <summary>
        /// Retrieve the next events created on the current site and all subsites
        /// </summary>
        /// <returns></returns>
        IEnumerable<SearchEventItem> GetNextEvents();

        /// <summary>
        /// Execute search query on SharePoint
        /// </summary>
        /// <param name="query">Search query</param>
        /// <param name="sorting">Sorting properties</param>
        /// <param name="rankModel">Ranking model to use for searching</param>
        /// <returns></returns>
        ClientResult<ResultTableCollection> ExecuteSearchQuery(
            string query,
            string sorting = null,
            string rankModel = null);
    }
}
