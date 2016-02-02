using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SharePointAddInsSampleWeb.Service.Contracts;
using SharePointAddInsSampleWeb.Service.Contracts.Data;

namespace SharePointAddInsSampleWeb.Service.Impl
{
    public class SearchService : ISearchService
    {
        /// <summary>
        /// SharePoint Client context
        /// </summary>
        private readonly ClientContext _clientContext;

        /// <summary>
        /// Search service constructor
        /// </summary>
        /// <param name="context"></param>
        public SearchService(ClientContext context)
        {
            _clientContext = context;
        }


        public IEnumerable<SearchDocumentItem> GetLatestModifiedDocuments()
        {
                // Retrieve current web url
                _clientContext.Load(_clientContext.Web, w => w.Url);
                _clientContext.ExecuteQuery();
                var webUrl = _clientContext.Web.Url;

               var query = $"Path:\"{webUrl}\" (IsDocument:True) (FileExtension:doc OR FileExtension:docx OR " +
                           "FileExtension:ppt OR FileExtension:pptx OR FileExtension:xls OR " +
                           "FileExtension:xlsx OR FileExtension:pdf)";

                var results = ExecuteSearchQuery(query);
                _clientContext.ExecuteQuery();
                var searchResults = results.Value[0].ResultRows.Select(MapResultRowToSearchDocumentItem).ToList();
                return searchResults.OrderByDescending(s => s.LastModifiedDate);
        }


        public IEnumerable<SearchDocumentItem> GetMostPopularDocuments()
        {
            // Retrieve current web url
            _clientContext.Load(_clientContext.Web, w => w.Url);
            _clientContext.ExecuteQuery();
            var webUrl = _clientContext.Web.Url;

            var query = $"Path:\"{webUrl}\" (IsDocument:True) (FileExtension:doc OR FileExtension:docx OR " +
                        "FileExtension:ppt OR FileExtension:pptx OR FileExtension:xls OR " +
                        "FileExtension:xlsx OR FileExtension:pdf)";

            // Popularity ranking model: d4ac6500-d1d0-48aa-86d4-8fe9a57a74af
            // Ranking model for popularity based search. This ranking model ranks SharePoint content 
            // based on the number of times an item that is stored in SharePoint has been accessed.
            var results = ExecuteSearchQuery(query, "Rank", "d4ac6500-d1d0-48aa-86d4-8fe9a57a74af");
            _clientContext.ExecuteQuery();

            var searchResults = results.Value[0].ResultRows.Select(MapResultRowToSearchDocumentItem).ToList();
            return searchResults.OrderByDescending(s => s.Rank);
        }

        public IEnumerable<SearchWebItem> GetMostPopularProjects()
        {
            // Retrieve current web url
            _clientContext.Load(_clientContext.Web, w => w.Url);
            _clientContext.ExecuteQuery();
            var webUrl = _clientContext.Web.Url;

            var query = $"Path:\"{webUrl}\" ContentClass:STS_Web";

            // Popularity ranking model: d4ac6500-d1d0-48aa-86d4-8fe9a57a74af
            // Ranking model for popularity based search. This ranking model ranks SharePoint content 
            // based on the number of times an item that is stored in SharePoint has been accessed.
            var results = ExecuteSearchQuery(query, "Rank", "d4ac6500-d1d0-48aa-86d4-8fe9a57a74af");
            _clientContext.ExecuteQuery();

            return results.Value[0].ResultRows.Select(
                MapResultRowToSearchWebItem).Where(item => item != null).OrderByDescending(s => s.Rank).ToList();
        }

        public IEnumerable<SearchTaskItem> GetMyActiveTasks()
        {
            // Retrieve current web url
            _clientContext.Load(_clientContext.Web, w => w.Url, w => w.CurrentUser.Title);
            _clientContext.ExecuteQuery();
            var webUrl = _clientContext.Web.Url;
            var currentUser = _clientContext.Web.CurrentUser.Title;

            var query = $"Path:\"{webUrl}\" ContentClass:STS_ListItem_Tasks PercentCompleteOWSNMBR<>1  " +
                        $"AssignedTo=\"{currentUser}\"";

            var results = ExecuteSearchQuery(query);
            _clientContext.ExecuteQuery();

            var searchResults = results.Value[0].ResultRows.Select(MapResultRowToSearchTaskItem).ToList();
            return searchResults.OrderBy(s => s.DueDate);
        }

        public IEnumerable<SearchEventItem> GetNextEvents()
        {
            // Retrieve current web url
            _clientContext.Load(_clientContext.Web, w => w.Url, w => w.CurrentUser.Title);
            _clientContext.ExecuteQuery();
            var webUrl = _clientContext.Web.Url;

            var query = $"Path:\"{webUrl}\" ContentClass:STS_ListItem_Events";
            var results = ExecuteSearchQuery(query);
            _clientContext.ExecuteQuery();

            var searchResults = results.Value[0].ResultRows.Select(MapResultRowToSearchEventItem).ToList();
            return searchResults.Where(u => u.StartDate.HasValue && (u.StartDate.Value.Date >= DateTime.Now.Date)).OrderBy(s => s.StartDate);
        }

        public ClientResult<ResultTableCollection> ExecuteSearchQuery(
            string query, string sorting = null, string rankModel = null)
        {
            var keywordQuery = new KeywordQuery(_clientContext){ QueryText = query};

            // Ranking model to use
            if (!String.IsNullOrEmpty(rankModel))
                keywordQuery.RankingModelId = rankModel;

            // If sorting parameter provided
            if (!String.IsNullOrEmpty(sorting))
                keywordQuery.SortList.Add(sorting, SortDirection.Descending);

            // Select only this properties for documents
            keywordQuery.SelectProperties.Add("Title");
            keywordQuery.SelectProperties.Add("Path");
            keywordQuery.SelectProperties.Add("ServerRedirectedURL");
            keywordQuery.SelectProperties.Add("ModifiedOWSDATE");
            keywordQuery.SelectProperties.Add("ModifiedBy");
            keywordQuery.SelectProperties.Add("CreatedOWSDATE");
            keywordQuery.SelectProperties.Add("CreatedBy");
            keywordQuery.SelectProperties.Add("SPWebUrl");
            keywordQuery.SelectProperties.Add("Rank");
            keywordQuery.SelectProperties.Add("LastModifiedTime");

            // Select only this properties for tasks
            keywordQuery.SelectProperties.Add("StatusOWSCHCS");
            keywordQuery.SelectProperties.Add("PercentCompleteOWSNMBR");
            keywordQuery.SelectProperties.Add("AssignedTo");
            keywordQuery.SelectProperties.Add("DueDateOWSDATE");
            keywordQuery.SelectProperties.Add("StartDateOWSDATE");


            // Select only this properties for events
            keywordQuery.SelectProperties.Add("EventDateOWSDATE");
            keywordQuery.SelectProperties.Add("EndDateOWSDATE");

            var searchExecutor = new SearchExecutor(_clientContext);
            var results = searchExecutor.ExecuteQuery(keywordQuery);
            return results;
        }

        #region Mapping search items

        /// <summary>
        /// Map SharePoint results to search document item
        /// </summary>
        /// <param name="resultRow"></param>
        /// <returns></returns>
        private SearchDocumentItem MapResultRowToSearchDocumentItem(IDictionary<string, object> resultRow)
        {
            var searchItem = new SearchDocumentItem
            {
                Title = resultRow["Title"] as string,
                Link = resultRow["Path"] as string,
                OnlineLink = resultRow["ServerRedirectedURL"] != null
                    ? resultRow["ServerRedirectedURL"] as string
                    : "",
                ModifiedBy = resultRow["ModifiedBy"] as string,
                CreatedBy = resultRow["CreatedBy"] as string,
                WebUrl = resultRow["SPWebUrl"] as string,
            };


            float rank;
            if (float.TryParse(resultRow["Rank"].ToString(), out rank))
                searchItem.Rank = rank;

            DateTime lastModifiedDate;
            if (DateTime.TryParse(resultRow["ModifiedOWSDATE"].ToString(), out lastModifiedDate))
                searchItem.LastModifiedDate = lastModifiedDate;

            DateTime createdDate;
            if (DateTime.TryParse(resultRow["CreatedOWSDATE"].ToString(), out createdDate))
                searchItem.CreatedDate = createdDate;

            return searchItem;
        }

        /// <summary>
        /// Map SharePoint results to search web item
        /// </summary>
        /// <param name="resultRow"></param>
        /// <returns></returns>
        private SearchWebItem MapResultRowToSearchWebItem(IDictionary<string, object> resultRow)
        {
            var searchItem = new SearchWebItem
            {
                Title = resultRow["Title"] as string,
                Link = resultRow["Path"] as string,
            };

            DateTime lastModifiedDate;
            if (resultRow["LastModifiedTime"] != null &&
                DateTime.TryParse(resultRow["LastModifiedTime"].ToString(), out lastModifiedDate))
                searchItem.LastModifiedDate = lastModifiedDate;

            float rank;
            if (float.TryParse(resultRow["Rank"].ToString(), out rank))
                searchItem.Rank = rank;

            return searchItem;
        }

        /// <summary>
        /// Map SharePoint results to search task item
        /// </summary>
        /// <param name="resultRow"></param>
        /// <returns></returns>
        private SearchTaskItem MapResultRowToSearchTaskItem(IDictionary<string, object> resultRow)
        {
            var searchItem = new SearchTaskItem
            {
                Title = resultRow["Title"] as string,
                Link = resultRow["Path"] as string,
                Status = resultRow["StatusOWSCHCS"] != null ? resultRow["StatusOWSCHCS"] as string : "",
                AssignedTo = resultRow["AssignedTo"] != null ? resultRow["AssignedTo"] as string : "",
                WebUrl = resultRow["SPWebUrl"] as string,
            };

            float completion;
            if (resultRow["PercentCompleteOWSNMBR"] != null
                    && float.TryParse(resultRow["PercentCompleteOWSNMBR"].ToString(), out completion))
                searchItem.CompletionPercentage = completion;

            DateTime dueDate;
            if (resultRow["DueDateOWSDATE"] != null
                    && DateTime.TryParse(resultRow["DueDateOWSDATE"].ToString(), out dueDate))
                searchItem.DueDate = dueDate;

            DateTime startDate;
            if (resultRow["StartDateOWSDATE"] != null
                    && DateTime.TryParse(resultRow["StartDateOWSDATE"].ToString(), out startDate))
                searchItem.StartDate = startDate;

            return searchItem;
        }


        /// <summary>
        /// Map SharePoint results to search event item
        /// </summary>
        /// <param name="resultRow"></param>
        /// <returns></returns>
        private SearchEventItem MapResultRowToSearchEventItem(IDictionary<string, object> resultRow)
        {
            var searchItem = new SearchEventItem
            {
                Title = resultRow["Title"] as string,
                Link = resultRow["Path"] as string,
                WebUrl = resultRow["SPWebUrl"] as string,
            };

            DateTime startDate;
            if (resultRow["EventDateOWSDATE"] != null && DateTime.TryParse(resultRow["EventDateOWSDATE"].ToString(), out startDate))
                searchItem.StartDate = startDate;

            DateTime endDate;
            if (resultRow["EndDateOWSDATE"] != null && DateTime.TryParse(resultRow["EndDateOWSDATE"].ToString(), out endDate))
                searchItem.EndDate = endDate;

            return searchItem;
        }

        #endregion


    }
}
