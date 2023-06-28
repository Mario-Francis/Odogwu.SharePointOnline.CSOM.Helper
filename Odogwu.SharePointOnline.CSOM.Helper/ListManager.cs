using Microsoft.SharePoint.Client;
using Odogwu.SharePointOnline.CSOM.Helper.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public class ListManager : IDisposable, IListManager
    {
        private readonly AuthenticationManager authMgr;

        public ListManager(AuthenticationManager authMgr)
        {
            this.authMgr = authMgr;
        }

        public string SiteUrl
        {
            get
            {
                return authMgr?.SiteUrl.ToString();
            }
            set
            {
                authMgr.SiteUrl = new Uri(value);
            }
        }

        // add new list item
        public async Task<int> AddItem(AddItemRequest request)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.ListName);

                context.Load(list, l => l.Fields);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                string invalidField;
                var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.FieldValues.Select(p => p.Key), out invalidField);
                if (!isFieldsValid)
                    throw new Exception($"Invalid field name '{invalidField}'! Field does not exist or is readonly.");

                SP.ListItemCreationInformation listItemCreationInfo = new SP.ListItemCreationInformation();
                SP.ListItem item = list.AddItem(listItemCreationInfo);

                foreach(var f in request.FieldValues)
                {
                    item[f.Key] = f.Value;
                }
                item.Update();
                context.Load(item);

                await context.ExecuteQueryAsync();

                return item.Id;
            }
        }

        // add batch list items
        public async Task BatchAddItems(BatchAddItemsRequest request, int maxItemInBatch=100)
        {
            if (request?.Items.Count() > maxItemInBatch)
            {
                throw new Exception($"Maximum add item count of {maxItemInBatch} exceeded!");
            }

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.ListName);

                context.Load(list, l => l.Fields);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

               for(var i = 0; i < request.Items.Count(); i++)
                {
                    var newItemFields = request.Items.ElementAt(i);

                    string invalidField;
                    var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), newItemFields.Select(p => p.Key), out invalidField);
                    if (!isFieldsValid)
                        throw new Exception($"Invalid field name '{invalidField}' in item at index {i}! Field does not exist or is readonly.");

                    SP.ListItemCreationInformation listItemCreationInfo = new SP.ListItemCreationInformation();
                    SP.ListItem item = list.AddItem(listItemCreationInfo);

                    foreach (var f in newItemFields)
                    {
                        item[f.Key] = f.Value;
                    }
                    item.Update();
                }

                await context.ExecuteQueryAsync();
            }
        }
        // update list item
        public async Task UpdateItem(UpdateItemRequest request)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.ListName);

                context.Load(list, l => l.Fields);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                string invalidField;
                var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.FieldValues.Select(p => p.Key), out invalidField);
                if (!isFieldsValid)
                    throw new Exception($"Invalid field name '{invalidField}'! Field does not exist or is readonly.");

                var item = list.GetItemById(request.Id);
                context.Load(item);

                foreach (var f in request.FieldValues)
                {
                    item[f.Key] = f.Value;
                }
                item.Update();
                context.Load(item);

                await context.ExecuteQueryAsync();
            }
        }

        // delete list item
        public async Task DeleteItem(int id, string listName)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(listName);

                var item = list.GetItemById(id);

                context.Load(item);
                item.DeleteObject();
                await context.ExecuteQueryAsync();
            }
        }
        // search by column value
        public async Task<ListSearchResult> SearchListByFieldValues(SearchListByFieldValuesRequest request, int maxResultLength = 100, int maxItemLoad = 500000)
        {
            var res = new ListSearchResult();
            var items = new List<SPListItem>();

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = context.Web.Lists.GetByTitle(request.ListName);

                context.Load(list, l => l, l => l.Fields);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                string invalidField;
                var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.SearchParams.Select(p => p.Key), out invalidField);
                if (!isFieldsValid)
                    throw new Exception($"Invalid field name '{invalidField}'! Field does not exist or is readonly.");

                var andTree = Utilities.GenerateCAMLAndTree(request.SearchParams.Sanitize().ToList());
                var orderByDateCreatedXml = "";
                if (request.Options.SortByDateCreated)
                {
                    orderByDateCreatedXml = Utilities.GetOrderByDateCreatedXml(request.Options.SortByDateCreatedDir);
                }
                SP.CamlQuery query = new SP.CamlQuery()
                {
                    ViewXml = $"<View Scope='Recursive'><Query><Where>{andTree}</Where>{orderByDateCreatedXml}</Query><RowLimit>5000</RowLimit></View>"
                };
                var resultList = new List<SP.ListItem>();
                do
                {
                    SP.ListItemCollection listItems = list.GetItems(query);
                    context.Load(listItems, t => t.Include(t => null), t => t.ListItemCollectionPosition);
                    await context.ExecuteQueryAsync();

                    resultList.AddRange(listItems);
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null && resultList.Count < maxItemLoad);

                var total = resultList.Count;
                var displayList = resultList.Skip(request.Options.StartIndex).Take(request.Options.Length > maxResultLength ? maxResultLength : request.Options.Length);

                foreach (var item in displayList)
                {
                    var _listItem = new SPListItem();
                    context.Load(item, item=>item);
                    await context.ExecuteQueryAsync();

                    List<KeyValuePair<string, object>> fieldValues = new List<KeyValuePair<string, object>>();

                    fieldValues.Add(new KeyValuePair<string, object>("ID", item.FieldValues["ID"]?.ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Title", item.FieldValues["Title"]?.ToString()));
                    foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                    {
                        var name = lf.InternalName.Replace("_x0020_", "_");
                        fieldValues.Add(new KeyValuePair<string, object>(name, item.FieldValues[lf.InternalName]));
                    }
                   
                    fieldValues.Add(new KeyValuePair<string, object>("Created", item.FieldValues["Created"].ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Modified", item.FieldValues["Modified"].ToString()));

                    var author = item.FieldValues["Author"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                    var editor = item.FieldValues["Editor"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                    _listItem.Id = item.Id;
                    _listItem.FieldValues = fieldValues;

                    items.Add(_listItem);
                }
                // return file
                res.ListItems = items;
                res.TotalResultCount = total;
                res.TotalDisplayCount = items.Count;

                return res;
            }
        }

        // search  by date range
        public async Task<ListSearchResult> SearchListByDateRange(SearchListByDateRangeRequest request, int maxResultLength = 100, int maxItemLoad = 500000)
        {
            var res = new ListSearchResult();
            var items = new List<SPListItem>();

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = context.Web.Lists.GetByTitle(request.ListName);

                var listFields = list.Fields;
                context.Load(listFields);

                var orderByDateCreatedXml = "";
                if (request.Options.SortByDateCreated)
                {
                    orderByDateCreatedXml = Utilities.GetOrderByDateCreatedXml(request.Options.SortByDateCreatedDir);
                }
                SP.CamlQuery query = new SP.CamlQuery()
                {
                    ViewXml = "<View Scope='Recursive'><Query>" +
                     "<Where><And>" +
                     $"<Geq><FieldRef Name='Created'/><Value Type='DateTime'>{request.StartDate.ToString("yyyy-MM-ddThh:mm:ss")}</Value></Geq>" +
                     $"<Leq><FieldRef Name='Created'/><Value Type='DateTime'>{request.EndDate.ToString("yyyy-MM-ddThh:mm:ss")}</Value></Leq>" +
                     "</And></Where>" + orderByDateCreatedXml +
                     "</Query><RowLimit>5000</RowLimit></View>"
                };
                var resultList = new List<SP.ListItem>();
                do
                {
                    SP.ListItemCollection listItems = list.GetItems(query);
                    context.Load(listItems, t => t.Include(t => null), t => t.ListItemCollectionPosition);
                    await context.ExecuteQueryAsync();

                    resultList.AddRange(listItems);
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null && resultList.Count < maxItemLoad);

                var total = resultList.Count;
                var displayList = resultList.Skip(request.Options.StartIndex).Take(request.Options.Length > maxResultLength ? maxResultLength : request.Options.Length);

                foreach (var item in displayList)
                {
                    var _listItem = new SPListItem();
                    context.Load(item, item => item);
                    await context.ExecuteQueryAsync();

                    List<KeyValuePair<string, object>> fieldValues = new List<KeyValuePair<string, object>>();

                    fieldValues.Add(new KeyValuePair<string, object>("ID", item.FieldValues["ID"]?.ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Title", item.FieldValues["Title"]?.ToString()));
                    foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                    {
                        var name = lf.InternalName.Replace("_x0020_", "_");
                        fieldValues.Add(new KeyValuePair<string, object>(name, item.FieldValues[lf.InternalName]));
                    }
                    fieldValues.Add(new KeyValuePair<string, object>("Created", item.FieldValues["Created"].ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Modified", item.FieldValues["Modified"].ToString()));

                    var author = item.FieldValues["Author"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                    var editor = item.FieldValues["Editor"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                    _listItem.Id = item.Id;
                    _listItem.FieldValues = fieldValues;

                    items.Add(_listItem);
                }
                // return file
                res.ListItems = items;
                res.TotalResultCount = total;
                res.TotalDisplayCount = items.Count;

                return res;
            }
        }

        // get by id
        public async Task<SPListItem> GetItem(int id, string listName)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(listName);

                var listFields = list.Fields;
                context.Load(listFields);

                var item = list.GetItemById(id);

                var _listItem = new SPListItem();
                context.Load(item, item => item);
                await context.ExecuteQueryAsync();

                List<KeyValuePair<string, object>> fieldValues = new List<KeyValuePair<string, object>>();

                fieldValues.Add(new KeyValuePair<string, object>("ID", item.FieldValues["ID"].ToString()));
                fieldValues.Add(new KeyValuePair<string, object>("Title", item.FieldValues["Title"]?.ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    fieldValues.Add(new KeyValuePair<string, object>(name, item.FieldValues[lf.InternalName]));
                }
                fieldValues.Add(new KeyValuePair<string, object>("Created", item.FieldValues["Created"].ToString()));
                fieldValues.Add(new KeyValuePair<string, object>("Modified", item.FieldValues["Modified"].ToString()));

                var author = item.FieldValues["Author"] as FieldUserValue;
                fieldValues.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                fieldValues.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                var editor = item.FieldValues["Editor"] as FieldUserValue;
                fieldValues.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                fieldValues.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                _listItem.Id = item.Id;
                _listItem.FieldValues =fieldValues;

                return _listItem;
            }
        }
        // get all list items
        public async Task<ListSearchResult> GetItems(GetListItemsRequest request, int maxResultLength = 100, int maxItemLoad = 500000)
        {
            var res = new ListSearchResult();
            var items = new List<SPListItem>();

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = context.Web.Lists.GetByTitle(request.ListName);

                var listFields = list.Fields;
                context.Load(listFields);

                var orderByDateCreatedXml = "";
                if (request.Options.SortByDateCreated)
                {
                    orderByDateCreatedXml = Utilities.GetOrderByDateCreatedXml(request.Options.SortByDateCreatedDir);
                }
                SP.CamlQuery query = new SP.CamlQuery()
                {
                    ViewXml = "<View Scope='Recursive'><Query>" + orderByDateCreatedXml +
                     "</Query><RowLimit>5000</RowLimit></View>"
                };
                var resultList = new List<SP.ListItem>();
                do
                {
                    SP.ListItemCollection listItems = list.GetItems(query);
                    context.Load(listItems, t => t.Include(t => null), t => t.ListItemCollectionPosition);
                    await context.ExecuteQueryAsync();

                    resultList.AddRange(listItems);
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null && resultList.Count < maxItemLoad);

                var total = resultList.Count;
                var displayList = resultList.Skip(request.Options.StartIndex).Take(request.Options.Length > maxResultLength ? maxResultLength : request.Options.Length);

                foreach(var l in displayList)
                {
                    context.Load(l, l => l);
                }
                await context.ExecuteQueryAsync();

                foreach (var item in displayList)
                {
                    var _listItem = new SPListItem();
                   // context.Load(item, item => item);
                   // await context.ExecuteQueryAsync();

                    List<KeyValuePair<string, object>> fieldValues = new List<KeyValuePair<string, object>>();

                    fieldValues.Add(new KeyValuePair<string, object>("ID", item.FieldValues["ID"].ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Title", item.FieldValues["Title"]?.ToString()));
                    foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                    {
                        var name = lf.InternalName.Replace("_x0020_", "_");
                        fieldValues.Add(new KeyValuePair<string, object>(name, item.FieldValues[lf.InternalName]));
                    }
                    fieldValues.Add(new KeyValuePair<string, object>("Created", item.FieldValues["Created"].ToString()));
                    fieldValues.Add(new KeyValuePair<string, object>("Modified", item.FieldValues["Modified"].ToString()));

                    var author = item.FieldValues["Author"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                    var editor = item.FieldValues["Editor"] as FieldUserValue;
                    fieldValues.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                    fieldValues.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                    _listItem.Id = item.Id;
                    _listItem.FieldValues =fieldValues;

                    items.Add(_listItem);
                }
                await context.ExecuteQueryAsync();
                // return file
                res.ListItems = items;
                res.TotalResultCount = total;
                res.TotalDisplayCount = items.Count;

                return res;
            }
        }
        // upload list item attachments
        public async Task<IEnumerable<ListItemAttachment>> UploadItemAttachments(int id, string listName, IEnumerable<AttachmentUploadItem> uploadItems)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(listName);

                var item = list.GetItemById(id);
                context.Load(item, item=>item, item=>item.AttachmentFiles);

                var responseList = new List<ListItemAttachment>();
                foreach (var uploadItem in uploadItems)
                {
                    var attInfo = new AttachmentCreationInformation();
                    attInfo.ContentStream = new MemoryStream(uploadItem.File);
                    attInfo.FileName = uploadItem.GetUniqueFileNameWithExtension();
                    var attachment = item.AttachmentFiles.Add(attInfo);

                    context.Load(attachment);
                    await context.ExecuteQueryAsync();

                    var response = new ListItemAttachment
                    {
                        ItemId=item.Id,
                        FileName = attachment.FileName,
                        ServerRelativeUrl = attachment.ServerRelativeUrl,
                        AbsoluteUrl = context.GetServerBaseUrl() + attachment.ServerRelativeUrl
                    };
                    responseList.Add(response);
                }
                return responseList;
            }
        }

        // add item with attachents
        public async Task<AddItemWithAttachmentsResponse> AddItemWithAttachments(AddItemWithAttachmentsRequest request)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.ListName);

                context.Load(list, l => l.Fields);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                string invalidField;
                var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.ItemFieldValues.Select(p => p.Key), out invalidField);
                if (!isFieldsValid)
                    throw new Exception($"Invalid field name '{invalidField}'! Field does not exist or is readonly.");

                SP.ListItemCreationInformation listItemCreationInfo = new SP.ListItemCreationInformation();
                SP.ListItem item = list.AddItem(listItemCreationInfo);

                foreach (var f in request.ItemFieldValues)
                {
                    item[f.Key] = f.Value;
                }
                item.Update();
                context.Load(item);
                await context.ExecuteQueryAsync();

                var responseList = new List<ListItemAttachment>();
                foreach (var uploadItem in request.UploadItems)
                {
                    var attInfo = new AttachmentCreationInformation();
                    attInfo.ContentStream = new MemoryStream(uploadItem.File);
                    attInfo.FileName = uploadItem.GetUniqueFileNameWithExtension();
                    var attachment = item.AttachmentFiles.Add(attInfo);

                    context.Load(attachment);
                    await context.ExecuteQueryAsync();

                    var response = new ListItemAttachment
                    {
                        ItemId=item.Id,
                        FileName = attachment.FileName,
                        ServerRelativeUrl = attachment.ServerRelativeUrl,
                        AbsoluteUrl = context.GetServerBaseUrl() + attachment.ServerRelativeUrl
                    };
                    responseList.Add(response);
                }

                return new AddItemWithAttachmentsResponse
                {
                    Id = item.Id,
                    Attachments = responseList
                };
            }
        }
        // get item with attachments
        public async Task<GetItemWithAttachmentsResponse> GetItemWithAttachments(int id, string listName)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(listName);

                var listFields = list.Fields;
                context.Load(listFields);

                var item = list.GetItemById(id);

                var _listItem = new SPListItem();
                context.Load(item, item => item, item=> item.AttachmentFiles);
                await context.ExecuteQueryAsync();

                List<KeyValuePair<string, object>> fieldValues = new List<KeyValuePair<string, object>>();

                fieldValues.Add(new KeyValuePair<string, object>("ID", item.FieldValues["ID"].ToString()));
                fieldValues.Add(new KeyValuePair<string, object>("Title", item.FieldValues["Title"]?.ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    fieldValues.Add(new KeyValuePair<string, object>(name, item.FieldValues[lf.InternalName]));
                }
                fieldValues.Add(new KeyValuePair<string, object>("Created", item.FieldValues["Created"].ToString()));
                fieldValues.Add(new KeyValuePair<string, object>("Modified", item.FieldValues["Modified"].ToString()));

                var author = item.FieldValues["Author"] as FieldUserValue;
                fieldValues.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                fieldValues.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                var editor = item.FieldValues["Editor"] as FieldUserValue;
                fieldValues.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                fieldValues.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                _listItem.Id = item.Id;
                _listItem.FieldValues =fieldValues;

                var attachments = new List<ListItemAttachment>();
                foreach(var attachment in item.AttachmentFiles)
                {
                    attachments.Add(new ListItemAttachment
                    {
                        FileName = attachment.FileName,
                        ItemId = item.Id,
                        ServerRelativeUrl = attachment.ServerRelativeUrl,
                        AbsoluteUrl = context.GetServerBaseUrl() + attachment.ServerRelativeUrl
                    });
                }

                return new GetItemWithAttachmentsResponse
                {
                    ListItem = _listItem,
                    Attachments = attachments
                };
            }
        }
        // get only attachments
        public async Task<IEnumerable<ListItemAttachment>> GetItemAttachments(int id, string listName)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(listName);

                var listFields = list.Fields;
                context.Load(listFields);

                var item = list.GetItemById(id);

                context.Load(item, item => item, item => item.AttachmentFiles);
                await context.ExecuteQueryAsync();

                var attachments = new List<ListItemAttachment>();
                foreach (var attachment in item.AttachmentFiles)
                {
                    attachments.Add(new ListItemAttachment
                    {
                        FileName = attachment.FileName,
                        ItemId = item.Id,
                        ServerRelativeUrl = attachment.ServerRelativeUrl,
                        AbsoluteUrl = context.GetServerBaseUrl() + attachment.ServerRelativeUrl
                    });
                }

                return attachments;
            }
        }

        public void Dispose()
        {
            authMgr.Dispose();
        }

    }
}
