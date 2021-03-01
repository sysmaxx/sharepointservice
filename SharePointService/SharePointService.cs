using Interfaces;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static Helper.SecurityHelper;
using static Helper.SharePointHelper;

namespace SharePoint
{
    public class SharePointService : IDisposable
    {
        private readonly ClientContext SharePointContext;

        public SharePointService(string url, string username, string password)
        {
            var _url = url ?? throw new ArgumentNullException(nameof(url));
            var _username = username ?? throw new ArgumentNullException(nameof(username));
            _ = password ?? throw new ArgumentNullException(nameof(password));
            
            var _clientCredentials = new SharePointOnlineCredentials(_username, ConvertToSecureString(password));
            SharePointContext = new ClientContext(_url)
            {
                Credentials = _clientCredentials
            };
        }

        #region Read List From SharePoint
        public List<T> ReadListFromSharePoint<T>(string collectionName)
            where T : IBaseModel, new ()
        {
            var selectedList = SharePointContext.Web.Lists.GetByTitle(collectionName);
            SharePointContext.Load(selectedList);
            CamlQuery camlQuery = new CamlQuery
            {
                ViewXml = $"<View></View>"
            };
            ListItemCollection sharePointItems = selectedList.GetItems(camlQuery);
            SharePointContext.Load(sharePointItems);
            SharePointContext.ExecuteQuery();

            return ParseSharePointList<T>(sharePointItems);
        }

        public async Task<List<T>> ReadListFromSharePointAsync<T>(string collectionName)
            where T : IBaseModel, new()
        {
            return await Task.Run(() => ReadListFromSharePoint<T>(collectionName)).ConfigureAwait(false);
        }


        #endregion

        #region Add Item To SharePoint Collection
        public int AddItemToSharePointList<T>(T item, string collectionName)
            where T : IBaseModel
        {
            var selectedList = SharePointContext.Web.Lists.GetByTitle(collectionName);
            ListItem newListItem = CreateListItem(item, selectedList);
            newListItem.Update();
            SharePointContext.ExecuteQuery();
            return newListItem.Id;
        }

        public async Task<int> AddItemToSharePointListAsync<T>(T item, string collectionName)
            where T : IBaseModel
        {
            return await Task.Run(() => AddItemToSharePointList(item, collectionName)).ConfigureAwait(false);
        }
        #endregion

        #region Add a Range of Items to SharePoint Collection
        public void AddRangeToSharePointList<T>(List<T> items, string collectionName)
            where T : IBaseModel
        {
            var selectedList = SharePointContext.Web.Lists.GetByTitle(collectionName);
            foreach (var item in items)
            {
                ListItem newListItem = CreateListItem(item, selectedList);
                newListItem.Update();
            }
            SharePointContext.ExecuteQuery();
        }

        public async Task AddRangeToSharePointListAsync<T>(List<T> items, string collectionName)
            where T : IBaseModel
        {
            await Task.Run(() => AddRangeToSharePointList(items, collectionName)).ConfigureAwait(false);
        }
        #endregion

        #region Delete an SharePoint Entry by ID
        public void DeleteEntryFromSharePointListById(int id, string collectionName)
        {
            var selectedList = SharePointContext.Web.Lists.GetByTitle(collectionName);
            var selectedItem = selectedList.GetItemById(id);

            selectedItem.DeleteObject();
            selectedList.Update();
            SharePointContext.ExecuteQuery();
        }

        public async Task DeleteEntryFromSharePointListByIdAsync<T>(int id, string collectionName)
        {
            await Task.Run(() => DeleteEntryFromSharePointListById(id, collectionName)).ConfigureAwait(false);
        }
        #endregion

        #region Update existing SharePoint Entry
        public void UpdateExistingItemOnSharePointList<T>(T item, string collectionName)
            where T : IBaseModel
        {
            if (!item.ID.HasValue)
                throw new NullReferenceException("ID cant be NULL");

            var selectedList = SharePointContext.Web.Lists.GetByTitle(collectionName);
            var selectedItem = selectedList.GetItemById(item.ID.Value);

            UpdateListItem(item, selectedItem);

            selectedItem.Update();
            SharePointContext.ExecuteQuery();
        }

        public async Task UpdateExistingItemOnSharePointListAsync<T>(T item, string collectionName)
            where T : IBaseModel
        {
            await Task.Run(() => UpdateExistingItemOnSharePointList(item, collectionName)).ConfigureAwait(false);
        }

        #endregion

        public void Dispose()
        {
            SharePointContext.Dispose();
        }
    }
}
