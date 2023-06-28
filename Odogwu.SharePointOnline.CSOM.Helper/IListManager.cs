using Odogwu.SharePointOnline.CSOM.Helper.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public interface IListManager
    {
        string SiteUrl { get; set; }
        Task<int> AddItem(AddItemRequest request);

        // add batch list items
        Task BatchAddItems(BatchAddItemsRequest request, int maxItemInBatch = 100);
        // update list item
        Task UpdateItem(UpdateItemRequest request);

        // delete list item
        Task DeleteItem(int id, string listName);
        // search by column value
        Task<ListSearchResult> SearchListByFieldValues(SearchListByFieldValuesRequest request, int maxResultLength = 100, int maxItemLoad = 500000);

        // search  by date range
        Task<ListSearchResult> SearchListByDateRange(SearchListByDateRangeRequest request, int maxResultLength = 100, int maxItemLoad = 500000);

        // get by id
        Task<SPListItem> GetItem(int id, string listName);
        // get all list items
        Task<ListSearchResult> GetItems(GetListItemsRequest request, int maxResultLength = 100, int maxItemLoad = 500000);
        // upload list item attachments
        Task<IEnumerable<ListItemAttachment>> UploadItemAttachments(int id, string listName, IEnumerable<AttachmentUploadItem> uploadItems);

        // add item with attachents
        Task<AddItemWithAttachmentsResponse> AddItemWithAttachments(AddItemWithAttachmentsRequest request);
        // get item with attachments
        Task<GetItemWithAttachmentsResponse> GetItemWithAttachments(int id, string listName);
        // get only attachments
        Task<IEnumerable<ListItemAttachment>> GetItemAttachments(int id, string listName);

    }
}
