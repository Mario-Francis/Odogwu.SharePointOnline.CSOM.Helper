using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Odogwu.SharePointOnline.CSOM.Helper.Models;
using SP = Microsoft.SharePoint.Client;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public class LibraryManager : IDisposable, ILibraryManager
    {
        private readonly AuthenticationManager authMgr;

        public LibraryManager(AuthenticationManager authMgr)
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

        /// <summary>
        /// Upload File
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public async Task<FileUploadResponse> UploadFile(FileUploadRequest request)
        {
            var res = new FileUploadResponse();
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);

                context.Load(list, l => l.Fields, l => l.RootFolder);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                string invalidField;
                var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.UploadItem.Properties.Select(p => p.Key), out invalidField);
                if (!isFieldsValid)
                    throw new Exception($"Invalid field name '{invalidField}'! Field does not exist or is readonly.");

                // process destination folder
                request.DestinationFolder = request.DestinationFolder ?? "";
                var destinationFolder = list.RootFolder;

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    destinationFolder = CreateFolder(destinationFolder, request.DestinationFolder);
                }
                else
                {
                    var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.DestinationFolder}";
                    if (!await IsFolderExists(web, destinationFolderUrl))
                    {
                        throw new Exception($"Folder with server relative url '{destinationFolderUrl}' does not exist");
                    }
                    destinationFolder = web.GetFolderByServerRelativeUrl(destinationFolderUrl);
                }

                //context.Load(destinationFolder.Files);
                // upload file
                var fci = new SP.FileCreationInformation();
                fci.ContentStream = new MemoryStream(request.UploadItem.File);
                fci.Url = request.UploadItem.GetUniqueFileNameWithExtension();
                fci.Overwrite = true;

                var file = destinationFolder.Files.Add(fci);
                context.Load(file);

                file.CheckOut();
                // set properties
                SP.ListItem item = file.ListItemAllFields;
                foreach (var p in request.UploadItem.Properties)
                {
                    item[p.Key] = p.Value;
                }
                item.Update();

                file.CheckIn(string.Empty, SP.CheckinType.OverwriteCheckIn);

                var _item = file.ListItemAllFields;
                context.Load(_item);

                await context.ExecuteQueryAsync();
                // return file info
                res.Id = _item.Id;
                res.Guid = file.UniqueId.ToString();
                res.FileName = file.Name;
                res.ServerRelativeUrl = file.ServerRelativeUrl;
                res.AbsoluteUrl = context.GetServerBaseUrl()+ file.ServerRelativeUrl;

                return res;
            }
        }
       
        /// <summary>
        /// Batch Upload File
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public async Task<IEnumerable<FileUploadResponse>> BatchUploadFile(BatchFileUploadRequest request, int maxUploadItem = 10)
        {
            if (request.UploadItems.Count() > maxUploadItem)
            {
                throw new Exception($"Maximum upload item count of {maxUploadItem} exceeded!");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);

                context.Load(list, l => l.Fields, l => l.RootFolder);
                await context.ExecuteQueryAsync();
                var listFields = list.Fields;

                // validate properties
                for (int i = 0; i < request.UploadItems.Count(); i++)
                {
                    string invalidField;
                    var isFieldsValid = Utilities.ValidateFieldNames(listFields.Where(lf => !lf.Hidden && lf.CanBeDeleted).Select(lf => lf.InternalName), request.UploadItems.ElementAtOrDefault(i).Properties.Select(p => p.Key), out invalidField);
                    if (!isFieldsValid)
                        throw new Exception($"Invalid field name '{invalidField}' for file at index {i}! Field does not exist or is readonly.");
                }

                // process destination folder
                request.DestinationFolder = request.DestinationFolder ?? "";
                var destinationFolder = list.RootFolder;

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    destinationFolder = CreateFolder(destinationFolder, request.DestinationFolder);
                }
                else
                {
                    var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.DestinationFolder}";
                    if (!await IsFolderExists(web, destinationFolderUrl))
                    {
                        throw new Exception($"Folder with server relative url '{destinationFolderUrl}' does not exist");
                    }
                    destinationFolder = web.GetFolderByServerRelativeUrl(destinationFolderUrl);
                }
                //context.Load(destinationFolder.Files);

                var responseList = new List<FileUploadResponse>();

                // upload files
                foreach (var uploadItem in request.UploadItems)
                {
                    var fci = new SP.FileCreationInformation();
                    fci.ContentStream = new MemoryStream(uploadItem.File);
                    fci.Url = uploadItem.GetUniqueFileNameWithExtension();
                    fci.Overwrite = true;

                    var file = destinationFolder.Files.Add(fci);
                    context.Load(file);

                    file.CheckOut();
                    // set properties
                    SP.ListItem item = file.ListItemAllFields;
                    foreach (var p in uploadItem.Properties)
                    {
                        item[p.Key] = p.Value;
                    }
                    item.Update();

                    file.CheckIn(string.Empty, SP.CheckinType.OverwriteCheckIn);

                    var _item = file.ListItemAllFields;
                    context.Load(_item);

                    await context.ExecuteQueryAsync();
                    // add to result list
                    var res = new FileUploadResponse
                    {
                        Id = _item.Id,
                        Guid = file.UniqueId.ToString(),
                        FileName = file.Name,
                        ServerRelativeUrl = file.ServerRelativeUrl,
                        AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl
                    };
                    responseList.Add(res);
                }
                return responseList;
            }
        }

        public async Task<SPFile> GetFileById(int id, string library)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                var list = context.Web.Lists.GetByTitle(library);
                var listFields = list.Fields;
                context.Load(listFields);

                var item = list.GetItemById(id);
                ConditionalScope scope = new ConditionalScope(context,() => item.FileSystemObjectType == FileSystemObjectType.File);
                ClientResult<Stream> clientResult = null;
                using (scope.StartScope()) {
                    context.Load(item, item => item, item => item.File);
                };

                await context.ExecuteQueryAsync();
                if (!scope.TestResult.Value)
                {
                    throw new Exception($"Invalid file id '{id}'");
                }
                else
                {
                    clientResult = item.File.OpenBinaryStream();
                    await context.ExecuteQueryAsync();
                }
                var file = item.File;
                var stream = clientResult.Value;
                MemoryStream ms = new MemoryStream();
                stream.CopyTo(ms);
                res.File = ms.ToArray();

                var itemFields = item.FieldValues;

                List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                }
                properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                var author = itemFields["Author"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                properties.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                var editor = itemFields["Editor"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                properties.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                // retirn file
                res.FileExtension = properties.FirstOrDefault(p => p.Key == "FileExtension").Value?.ToString();
                res.FileName = file.Name;
                res.Guid = file.UniqueId.ToString();
                res.Id = item.Id;
                res.Properties =properties;
                res.ServerRelativeUrl = file.ServerRelativeUrl;
                res.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                return res;
            }
        }

        public async Task<SPFile> GetFileByUniqueId(string uniqueId, string library)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;

                var list = context.Web.Lists.GetByTitle(library);
                var listFields = list.Fields;
                context.Load(listFields);

                var file = web.GetFileById(new Guid(uniqueId));
                context.Load(file);

                var item = file.ListItemAllFields;
                context.Load(item);

                var clientResult = file.OpenBinaryStream();
                await context.ExecuteQueryAsync();

                var stream = clientResult.Value;
                MemoryStream ms = new MemoryStream();
                stream.CopyTo(ms);
                res.File = ms.ToArray();

                var itemFields = item.FieldValues;

                List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                }
                properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                var author = itemFields["Author"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                properties.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                var editor = itemFields["Editor"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                properties.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                // retirn file
                res.FileExtension = properties.FirstOrDefault(p => p.Key == "FileExtension").Value?.ToString();
                res.FileName = file.Name;
                res.Guid = file.UniqueId.ToString();
                res.Id = item.Id;
                res.Properties =properties;
                res.ServerRelativeUrl = file.ServerRelativeUrl;
                res.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                return res;
            }

        }

        public async Task<SPFile> GetFileByUrl(string fileUrl, string library)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;

                var list = context.Web.Lists.GetByTitle(library);
                var listFields = list.Fields;
                context.Load(listFields);

                SP.File file;
                if (fileUrl.Contains("http"))
                {
                    file = context.Web.GetFileByUrl(fileUrl);
                }
                else
                {
                    file = context.Web.GetFileByServerRelativeUrl(fileUrl);
                }
                context.Load(file);

                var item = file.ListItemAllFields;
                context.Load(item);

                var clientResult = file.OpenBinaryStream();
                await context.ExecuteQueryAsync();

                var stream = clientResult.Value;
                MemoryStream ms = new MemoryStream();
                stream.CopyTo(ms);
                res.File = ms.ToArray();

                var itemFields = item.FieldValues;

                List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                }
                properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                var author = itemFields["Author"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                properties.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                var editor = itemFields["Editor"] as FieldUserValue;
                properties.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                properties.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));

                // return file
                res.FileExtension = properties.FirstOrDefault(p => p.Key == "FileExtension").Value?.ToString();
                res.FileName = file.Name;
                res.Guid = file.UniqueId.ToString();
                res.Id = item.Id;
                res.Properties =properties;
                res.ServerRelativeUrl = file.ServerRelativeUrl;
                res.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                return res;
            }
        }

        public async Task<FileSearchResult> SearchFilesByProperties(SearchFilesByPropertiesRequest request, int maxResultLength = 10, int maxItemLoad=500000)
        {
            var res = new FileSearchResult();
            var files = new List<SPFile>();

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = context.Web.Lists.GetByTitle(request.Library);

                context.Load(list, l=>l, l => l.Fields);
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
                var targetFolderUrl = context.ResolveTargetFolder(request.Library, request.TargetFolder);

                SP.CamlQuery query = new SP.CamlQuery()
                {
                    ViewXml = $"<View Scope='Recursive'><Query><Where>{andTree}</Where>{orderByDateCreatedXml}</Query><RowLimit>5000</RowLimit></View>",
                    FolderServerRelativeUrl = targetFolderUrl,
                    AllowIncrementalResults=true
                };

                var resultList = new List<ListItem>();
                do
                {
                    ListItemCollection listItems = list.GetItems(query);
                    context.Load(listItems, t => t.Include(t => null), t => t.ListItemCollectionPosition);
                    await context.ExecuteQueryAsync();

                    resultList.AddRange(listItems);
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null && resultList.Count < maxItemLoad);

                var total = resultList.Count;
                var displayList = resultList.Skip(request.Options.StartIndex).Take(request.Options.Length > maxResultLength ? maxResultLength : request.Options.Length);

                foreach (var item in displayList)
                {
                    var _file = new SPFile();

                    var file = item.File;
                    context.Load(file);

                    var _item = file.ListItemAllFields;
                    context.Load(_item);
                    var itemFields = _item.FieldValues;
                    await context.ExecuteQueryAsync();

                    List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                    if (request.IncludeProperties)
                    {
                        properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                        foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                        {
                            var name = lf.InternalName.Replace("_x0020_", "_");
                            properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                        }
                        properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                        var author = itemFields["Author"] as FieldUserValue;
                        properties.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                        properties.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                        var editor = itemFields["Editor"] as FieldUserValue;
                        properties.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                        properties.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));
                    }
                    _file.FileExtension = itemFields["File_x0020_Type"].ToString();
                    _file.FileName = file.Name;
                    _file.Guid = file.UniqueId.ToString();
                    _file.Id = _item.Id;
                    _file.Properties =properties;
                    _file.ServerRelativeUrl = file.ServerRelativeUrl;
                    _file.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                    files.Add(_file);
                }
                //await context.ExecuteQueryAsync();

                // return file
                res.Files = files;
                res.TotalResultCount = total;
                res.TotalDisplayCount = files.Count;

                return res;
            }
        }

        public async Task<FileSearchResult> SearchFilesByDateRange(SearchFilesByDateRangeRequest request, int maxResultLength = 10, int maxItemLoad = 500000)
        {
            var res = new FileSearchResult();
            var files = new List<SPFile>();

            using (var context = authMgr.GetContext())
            {
                var web = context.Web;

                var list = context.Web.Lists.GetByTitle(request.Library);
                var listFields = list.Fields;
                context.Load(listFields);

                var orderByDateCreatedXml = "";
                if (request.Options.SortByDateCreated)
                {
                    orderByDateCreatedXml = Utilities.GetOrderByDateCreatedXml(request.Options.SortByDateCreatedDir);
                }

                var targetFolderUrl = context.ResolveTargetFolder(request.Library, request.TargetFolder);

                SP.CamlQuery query = new SP.CamlQuery()
                {
                    ViewXml = "<View Scope='Recursive'><Query>" +
                     "<Where><And>" +
                     $"<Geq><FieldRef Name='Created'/><Value Type='DateTime'>{request.StartDate.ToString("yyyy-MM-ddThh:mm:ss")}</Value></Geq>" +
                     $"<Leq><FieldRef Name='Created'/><Value Type='DateTime'>{request.EndDate.ToString("yyyy-MM-ddThh:mm:ss")}</Value></Leq>" +
                     "</And></Where>" + orderByDateCreatedXml +
                     "</Query><RowLimit>5000</RowLimit></View>",
                    FolderServerRelativeUrl = targetFolderUrl,
                    AllowIncrementalResults=true
                };

                var resultList = new List<ListItem>();
                do
                {
                    ListItemCollection listItems = list.GetItems(query);
                    context.Load(listItems, t => t.Include(t => null), t => t.ListItemCollectionPosition);
                    await context.ExecuteQueryAsync();

                    resultList.AddRange(listItems);
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null && resultList.Count < maxItemLoad);

                var total = resultList.Count;
                var displayList = resultList.Skip(request.Options.StartIndex).Take(request.Options.Length > maxResultLength ? maxResultLength : request.Options.Length);

                foreach (var item in displayList)
                {
                    var _file = new SPFile();

                    var file = item.File;
                    context.Load(file);

                    var _item = file.ListItemAllFields;
                    context.Load(_item);
                    var itemFields = _item.FieldValues;
                    await context.ExecuteQueryAsync();

                    List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                    if (request.IncludeProperties)
                    {
                        properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                        foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                        {
                            var name = lf.InternalName.Replace("_x0020_", "_");
                            properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                        }
                        properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                        properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                        var author = itemFields["Author"] as FieldUserValue;
                        properties.Add(new KeyValuePair<string, object>("AuthorEmail", author.Email));
                        properties.Add(new KeyValuePair<string, object>("AuthorName", author.LookupValue));
                        var editor = itemFields["Editor"] as FieldUserValue;
                        properties.Add(new KeyValuePair<string, object>("EditorEmail", editor.Email));
                        properties.Add(new KeyValuePair<string, object>("EditorName", editor.LookupValue));
                    }
                    _file.FileExtension = itemFields["File_x0020_Type"].ToString();
                    _file.FileName = file.Name;
                    _file.Guid = file.UniqueId.ToString();
                    _file.Id = _item.Id;
                    _file.Properties =properties;
                    _file.ServerRelativeUrl = file.ServerRelativeUrl;
                    _file.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                    files.Add(_file);
                }
                //await context.ExecuteQueryAsync();

                // return file
                res.Files = files;
                res.TotalResultCount = total;
                res.TotalDisplayCount = files.Count;

                return res;
            }
        }

        public async Task<SPFile> UpdateFileProperties(UpdateFilePropertiesRequest request)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                var list = context.Web.Lists.GetByTitle(request.Library);
                var listFields = list.Fields;
                context.Load(listFields);

                var item = list.GetItemById(request.Id);
                ConditionalScope scope = new ConditionalScope(context, () => item.FileSystemObjectType == FileSystemObjectType.File);
                ClientResult<Stream> clientResult = null;
                using (scope.StartScope())
                {
                    context.Load(item, item => item, item => item.File);
                };

                await context.ExecuteQueryAsync();
                if (!scope.TestResult.Value)
                {
                    throw new Exception($"Invalid file id '{request.Id}'");
                }
                else
                {
                    clientResult = item.File.OpenBinaryStream();
                    await context.ExecuteQueryAsync();
                }
                var file = item.File;

                file.CheckOut();
                try
                {
                    foreach (var p in request.Properties)
                    {
                        item[p.Key] = p.Value;
                    }
                    item.Update();

                    // use OverwriteCheckIn type to make sure not to create multiple versions 
                    if (file.CheckOutType != CheckOutType.None)
                    {
                        file.CheckIn(string.Empty, SP.CheckinType.OverwriteCheckIn);
                    }
                    await context.ExecuteQueryAsync();
                }
                catch (SP.ServerException ex)
                {
                    throw ex;
                }
                finally
                {
                    file.CheckIn(string.Empty, SP.CheckinType.OverwriteCheckIn);
                    await context.ExecuteQueryAsync();
                }

                var itemFields = item.FieldValues;
                List<KeyValuePair<string, object>> properties = new List<KeyValuePair<string, object>>();
                properties.Add(new KeyValuePair<string, object>("ID", itemFields["ID"].ToString()));
                properties.Add(new KeyValuePair<string, object>("UniqueId", itemFields["UniqueId"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Name", itemFields["FileLeafRef"].ToString()));
                foreach (var lf in listFields.Where(f => !f.Hidden && f.CanBeDeleted))
                {
                    var name = lf.InternalName.Replace("_x0020_", "_");
                    properties.Add(new KeyValuePair<string, object>(name, itemFields[lf.InternalName]));
                }
                properties.Add(new KeyValuePair<string, object>("ServerRelativeUrl", itemFields["FileRef"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileExtension", itemFields["File_x0020_Type"].ToString()));
                properties.Add(new KeyValuePair<string, object>("FileSize", itemFields["File_x0020_Size"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Created", itemFields["Created"].ToString()));
                properties.Add(new KeyValuePair<string, object>("Modified", itemFields["Modified"].ToString()));

                // retirn file
                res.FileExtension = properties.FirstOrDefault(p => p.Key == "FileExtension").Value?.ToString();
                res.FileName = file.Name;
                res.Guid = file.UniqueId.ToString();
                res.Id = item.Id;
                res.Properties =properties;
                res.ServerRelativeUrl = file.ServerRelativeUrl;
                res.AbsoluteUrl = context.GetServerBaseUrl() + file.ServerRelativeUrl;

                return res;
            }



        }

        public async Task DeleteFileById(int id, string library)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                var list = context.Web.Lists.GetByTitle(library);

                var item = list.GetItemById(id);
                ConditionalScope scope = new ConditionalScope(context, () => item.FileSystemObjectType == FileSystemObjectType.File);
                using (scope.StartScope())
                {
                    context.Load(item, item => item, item => item.File);
                };

                await context.ExecuteQueryAsync();
                if (!scope.TestResult.Value)
                {
                    throw new Exception($"Invalid file id '{id}'");
                }

                var file = item.File;

                file.DeleteObject();
                await context.ExecuteQueryAsync();
            }
        }
        public async Task DeleteFileByUrl(string fileUrl)
        {
            var res = new SPFile();
            using (var context = authMgr.GetContext())
            {
                SP.File file;
                if (fileUrl.Contains("http"))
                {
                    file = context.Web.GetFileByUrl(fileUrl);
                }
                else
                {
                    file = context.Web.GetFileByServerRelativeUrl(fileUrl);
                }
                context.Load(file);

                file.DeleteObject();
                await context.ExecuteQueryAsync();
            }
        }
        
        /// <summary>
        /// Delete folder
        /// </summary>
        /// <param name="libraryRelativeFolderPath">Path to the folder relative to the document library</param>
        /// <param name="library">Library name</param>
        /// <returns></returns>
        public async Task DeleteFolder(string libraryRelativeFolderPath, string library)
        {
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                libraryRelativeFolderPath = libraryRelativeFolderPath?.Replace("../", "")?.Replace("./", "").Replace("\\", "")?.Trim('/')?.Trim('.').Trim();
                if (string.IsNullOrEmpty(libraryRelativeFolderPath))
                {
                    throw new Exception("Library relative folder path is required");
                }

                var folderSiteRelativeUrl = $"{context.GetSiteServerRelativeUrl()}/{library}/{libraryRelativeFolderPath}";
                if (!await IsFolderExists(web, folderSiteRelativeUrl))
                {
                    throw new Exception($"Folder with server relative url '{folderSiteRelativeUrl}' does not exist");
                }

                var folder = web.GetFolderByServerRelativeUrl(folderSiteRelativeUrl);
                folder.DeleteObject();
                await context.ExecuteQueryAsync();
            }
        }

        
        public async Task CopyFile(CopyFileRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFileUrl))
            {
                throw new Exception("Source file url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                if (!await IsFileExists(web, request.SourceFileUrl))
                {
                    throw new Exception($"File with server relative url '{request.SourceFileUrl}' does not exist");
                }
                var sourceFileName = Path.GetFileName(request.SourceFileUrl);
                var sourceExtension = Path.GetExtension(request.SourceFileUrl);
                var destinationFileName = $"{(string.IsNullOrEmpty(request.NewFileName?.Trim()) ? sourceFileName : (request.NewFileName?.Split(".")[0].Replace(" ", "_")+sourceExtension))}";
                var destinationFolder = $"{CleanFolderPath(request.DestinationFolder)}";
                var destinationFileUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{destinationFolder}/{destinationFileName}".Replace("//", "/");

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationFolder);
                }

                var overwrite = request.CopyFileOption == MoveCopyFileOptions.OverwriteDuplicate;
                var rename = request.CopyFileOption == MoveCopyFileOptions.RenameDuplicate;
                if (request.CopyFileOption == MoveCopyFileOptions.ReportDuplicate)
                {
                    if (await IsFileExists(web, destinationFileUrl))
                    {
                        throw new Exception($"File with name '{destinationFileName}' already exists in the specified destination folder '{request.DestinationFolder}'");
                    }
                }
                var sourceFileUrl = request.SourceFileUrl.StartsWith("http") ? request.SourceFileUrl : context.GetServerBaseUrl() + "/" + request.SourceFileUrl.TrimStart('/');
                destinationFileUrl = context.GetServerBaseUrl() + "/" + destinationFileUrl.TrimStart('/');

                MoveCopyUtil.CopyFile(context, sourceFileUrl, destinationFileUrl, overwrite, new MoveCopyOptions { KeepBoth = rename });
                await context.ExecuteQueryAsync();
            }
        }

        public async Task CopyFolder(CopyFolderRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFolderUrl?.Trim()))
            {
                throw new Exception("Library related source folder url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                var sourceFolderName = request.SourceFolderUrl.Split('/').Where(f => !string.IsNullOrEmpty(f)).LastOrDefault();
                if (string.IsNullOrEmpty(sourceFolderName?.Trim()))
                {
                    throw new Exception("Library related source folder url is required");
                }
                var sourceFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.SourceFolderUrl.TrimStart('/')}";
                var destinationParentFolder = $"{CleanFolderPath(request.DestinationFolder)}";
                var destinationFolderName = string.IsNullOrEmpty(request.NewFolderName) ? sourceFolderName : request.NewFolderName;
                var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.DestinationFolder.TrimStart('/')}/{destinationFolderName}";
                if (!await IsFolderExists(web, sourceFolderUrl))
                {
                    throw new Exception($"Folder with server relative url '{sourceFolderUrl}' does not exist");
                }

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationParentFolder);
                }

                var rename = request.CopyFolderOption == MoveCopyFolderOptions.RenameDuplicate;
                if (request.CopyFolderOption == MoveCopyFolderOptions.ReportDuplicate)
                {
                    if (await IsFolderExists(web, destinationFolderUrl))
                    {
                        throw new Exception($"Folder with name '{destinationFolderName}' already exists in the specified destination folder '{request.DestinationFolder}'");
                    }
                }
                var _sourceFolderUrl = context.GetServerBaseUrl() + "/" + sourceFolderUrl.TrimStart('/');
                var _destinationFolderUrl = context.GetServerBaseUrl() + "/" + destinationFolderUrl.TrimStart('/');

                MoveCopyUtil.CopyFolder(context, _sourceFolderUrl, _destinationFolderUrl, new MoveCopyOptions { KeepBoth = rename });
                await context.ExecuteQueryAsync();
            }
        }
        public async Task MoveFile(MoveFileRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFileUrl))
            {
                throw new Exception("Source file url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                if (!await IsFileExists(web, request.SourceFileUrl))
                {
                    throw new Exception($"File with server relative url '{request.SourceFileUrl}' does not exist");
                }
                var sourceFileName = Path.GetFileName(request.SourceFileUrl);
                var sourceExtension = Path.GetExtension(request.SourceFileUrl);
                var destinationFileName = $"{sourceFileName}";
                var destinationFolder = $"{CleanFolderPath(request.DestinationFolder)}";
                var destinationFileUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{destinationFolder}/{destinationFileName}".Replace("//", "/");

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationFolder);
                }

                var overwrite = request.MoveFileOption == MoveCopyFileOptions.OverwriteDuplicate;
                var rename = request.MoveFileOption == MoveCopyFileOptions.RenameDuplicate;
                if (request.MoveFileOption == MoveCopyFileOptions.ReportDuplicate)
                {
                    if (await IsFileExists(web, destinationFileUrl))
                    {
                        throw new Exception($"File with name '{destinationFileName}' already exists in the specified destination folder '{request.DestinationFolder}'");
                    }
                }
                var sourceFileUrl = request.SourceFileUrl.StartsWith("http") ? request.SourceFileUrl : context.GetServerBaseUrl() + "/" + request.SourceFileUrl.TrimStart('/');
                destinationFileUrl = context.GetServerBaseUrl() + "/" + destinationFileUrl.TrimStart('/');

                MoveCopyUtil.MoveFile(context, sourceFileUrl, destinationFileUrl, overwrite, new MoveCopyOptions { KeepBoth = rename });
                await context.ExecuteQueryAsync();
            }
        }

        public async Task MoveFolder(MoveFolderRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFolderUrl?.Trim()))
            {
                throw new Exception("Library related source folder url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                var sourceFolderName = request.SourceFolderUrl.Split('/').Where(f => !string.IsNullOrEmpty(f)).LastOrDefault();
                var sourceFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.SourceFolderUrl.TrimStart('/')}";
                var destinationParentFolder = $"{CleanFolderPath(request.DestinationFolder)}";
                var destinationFolderName = sourceFolderName;
                var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.DestinationFolder.TrimStart('/')}/{destinationFolderName}";
                if (!await IsFolderExists(web, sourceFolderUrl))
                {
                    throw new Exception($"Folder with server relative url '{sourceFolderUrl}' does not exist");
                }

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationParentFolder);
                }

                var rename = request.MoveFolderOption == MoveCopyFolderOptions.RenameDuplicate;
                if (request.MoveFolderOption == MoveCopyFolderOptions.ReportDuplicate)
                {
                    if (await IsFolderExists(web, destinationFolderUrl))
                    {
                        throw new Exception($"Folder with name '{destinationFolderName}' already exists in the specified destination folder '{request.DestinationFolder}'");
                    }
                }
                var _sourceFolderUrl = context.GetServerBaseUrl() + "/" + sourceFolderUrl.TrimStart('/');
                var _destinationFolderUrl = context.GetServerBaseUrl() + "/" + destinationFolderUrl.TrimStart('/');

                MoveCopyUtil.MoveFolder(context, _sourceFolderUrl, _destinationFolderUrl, new MoveCopyOptions { KeepBoth = rename });
                await context.ExecuteQueryAsync();
            }
        }

        public async Task<IEnumerable<string>> CopyFolderContents(MoveCopyFolderContentsRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFolder?.Trim()))
            {
                throw new Exception("Library related source folder url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                var sourceFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.SourceFolder.TrimStart('/')}";
                var destinationRelativeUrl =  CleanFolderPath(request.DestinationFolder);
                var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{destinationRelativeUrl}";
                if (!await IsFolderExists(web, sourceFolderUrl))
                {
                    throw new Exception($"Source folder with server relative url '{sourceFolderUrl}' does not exist");
                }

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationRelativeUrl);
                }

                if (!await IsFolderExists(web, destinationFolderUrl))
                {
                    throw new Exception($"Destination folder with server relative url '{destinationFolderUrl}' does not exist");
                }

                Folder sourceFolder = web.GetFolderByServerRelativeUrl(sourceFolderUrl);
                Folder destinationFolder = web.GetFolderByServerRelativeUrl(destinationFolderUrl);
                context.Load(sourceFolder, f => f.Files, f => f.Folders, f => f.ServerRelativeUrl);
                context.Load(destinationFolder, f => f.ServerRelativeUrl);
                await context.ExecuteQueryAsync();
                var errorMessages = new List<string>();
                if (request.MoveCopyContentType == FolderContentTypes.All || request.MoveCopyContentType == FolderContentTypes.FilesOnly)
                {
                    foreach (var f in sourceFolder.Files)
                    {
                        if (request.MoveCopyFolderContentOption == MoveCopyFolderContentOptions.RenameDuplicate)
                        {
                            MoveCopyUtil.CopyFile(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", false, new MoveCopyOptions { KeepBoth = true });
                        }
                        else
                        {
                            if (!await IsFileExists(web, f.ServerRelativeUrl.Replace(sourceFolder.ServerRelativeUrl, destinationFolder.ServerRelativeUrl)))
                            {
                                //f.CopyTo($"{destinationFolderUrl}/{f.Name}", false);
                                MoveCopyUtil.CopyFile(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", false, new MoveCopyOptions { KeepBoth = false });
                            }
                            else
                            {
                                errorMessages.Add($"File with url '{f.ServerRelativeUrl}' already exists in destination folder");
                            }
                        }
                    }
                }

                if (request.MoveCopyContentType == FolderContentTypes.All || request.MoveCopyContentType == FolderContentTypes.FoldersOnly)
                {
                    foreach (var f in sourceFolder.Folders)
                    {
                        if (request.MoveCopyFolderContentOption == MoveCopyFolderContentOptions.RenameDuplicate)
                        {
                            MoveCopyUtil.CopyFolder(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", new MoveCopyOptions { KeepBoth = true });
                        }
                        else
                        {
                            if (!await IsFolderExists(web, f.ServerRelativeUrl.Replace(sourceFolder.ServerRelativeUrl, destinationFolder.ServerRelativeUrl)))
                            {
                                //f.CopyTo($"{destinationFolderUrl}/{f.Name}");
                                MoveCopyUtil.CopyFolder(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", new MoveCopyOptions { KeepBoth = false });
                            }
                            else
                            {
                                errorMessages.Add($"Folder with url '{f.ServerRelativeUrl}' already exists in destination folder");
                            }
                        }
                    }
                }
                await context.ExecuteQueryAsync();
                return errorMessages;
            }
        }
        public async Task<IEnumerable<string>> MoveFolderContents(MoveCopyFolderContentsRequest request)
        {
            if (string.IsNullOrEmpty(request.SourceFolder?.Trim()))
            {
                throw new Exception("Library related source folder url is required");
            }
            request.CreateDestinationFolderIfNotExist = request.CreateDestinationFolderIfNotExist ?? true;
            using (var context = authMgr.GetContext())
            {
                var web = context.Web;
                var list = web.Lists.GetByTitle(request.Library);
                context.Load(list, l => l.RootFolder);

                var sourceFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{request.SourceFolder.TrimStart('/')}";
                var destinationRelativeUrl = CleanFolderPath(request.DestinationFolder);
                var destinationFolderUrl = $"{context.GetSiteServerRelativeUrl()}/{request.Library}/{destinationRelativeUrl}";
                if (!await IsFolderExists(web, sourceFolderUrl))
                {
                    throw new Exception($"Source folder with server relative url '{sourceFolderUrl}' does not exist");
                }

                if (request.CreateDestinationFolderIfNotExist.Value)
                {
                    CreateFolder(list.RootFolder, destinationRelativeUrl);
                }

                if (!await IsFolderExists(web, destinationFolderUrl))
                {
                    throw new Exception($"Destination folder with server relative url '{destinationFolderUrl}' does not exist");
                }

                Folder sourceFolder = web.GetFolderByServerRelativeUrl(sourceFolderUrl);
                Folder destinationFolder = web.GetFolderByServerRelativeUrl(destinationFolderUrl);
                context.Load(sourceFolder, f => f.Files, f => f.Folders, f => f.ServerRelativeUrl);
                context.Load(destinationFolder, f => f.ServerRelativeUrl);
                await context.ExecuteQueryAsync();
                var errorMessages = new List<string>();
                if (request.MoveCopyContentType == FolderContentTypes.All || request.MoveCopyContentType == FolderContentTypes.FilesOnly)
                {
                    foreach (var f in sourceFolder.Files)
                    {
                        if (request.MoveCopyFolderContentOption == MoveCopyFolderContentOptions.RenameDuplicate)
                        {
                            MoveCopyUtil.MoveFile(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", false, new MoveCopyOptions { KeepBoth = true });
                        }
                        else
                        {
                            if (!await IsFileExists(web, f.ServerRelativeUrl.Replace(sourceFolder.ServerRelativeUrl, destinationFolder.ServerRelativeUrl)))
                            {
                                //f.MoveTo($"{destinationFolderUrl}/{f.Name}", MoveOperations.AllowBrokenThickets);
                                MoveCopyUtil.MoveFile(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", false, new MoveCopyOptions { KeepBoth = false });
                            }
                            else
                            {
                                errorMessages.Add($"File with url '{f.ServerRelativeUrl}' already exists in destination folder");
                            }
                        }
                    }
                }

                if (request.MoveCopyContentType == FolderContentTypes.All || request.MoveCopyContentType == FolderContentTypes.FoldersOnly)
                {
                    foreach (var f in sourceFolder.Folders)
                    {
                        if (request.MoveCopyFolderContentOption == MoveCopyFolderContentOptions.RenameDuplicate)
                        {
                            MoveCopyUtil.MoveFolder(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", new MoveCopyOptions { KeepBoth = true });
                        }
                        else
                        {
                            if (!await IsFolderExists(web, f.ServerRelativeUrl.Replace(sourceFolder.ServerRelativeUrl, destinationFolder.ServerRelativeUrl)))
                            {
                                //f.MoveTo($"{destinationFolderUrl}/{f.Name}");
                                MoveCopyUtil.MoveFolder(context, $"{context.GetServerBaseUrl()}{f.ServerRelativeUrl}", $"{context.GetServerBaseUrl()}{destinationFolderUrl}/{f.Name}", new MoveCopyOptions { KeepBoth = false });
                            }
                            else
                            {
                                errorMessages.Add($"Folder with url '{f.ServerRelativeUrl}' already exists in destination folder");
                            }
                        }
                    }
                }
                await context.ExecuteQueryAsync();
                return errorMessages;
            }
        }

       
        public void Dispose()
        {
            authMgr.Dispose();
        }


        // private methods
        private string CleanFolderPath(string folderName)
        {
            var arr = folderName?.Split('/').Where(f => !string.IsNullOrEmpty(f)).Select(f => f.Trim('.').Replace("\\", ""));
            return string.Join('/', arr)?.Trim('/');
        }

        private Folder CreateFolder(Folder rootFolder, string folderPath)
        {
            Folder createdFolder = rootFolder;
            var destinationFolderHierarchy = folderPath.Trim().Split('/').Where(f => !string.IsNullOrEmpty(f)).Select(f => f.Replace(" ", "_").Replace(".", "_").Replace(",", ""));
            if (destinationFolderHierarchy.Count() != 0)
            {
                foreach (var folder in destinationFolderHierarchy)
                {
                    createdFolder = createdFolder.Folders.Add(folder);
                }
            }
            //await rootFolder.Context.ExecuteQueryAsync();
            return createdFolder;
        }

        private async Task<bool> IsFolderExists(Web web, string folderServerRelativeUrl)
        {
            try
            {
                var folder = web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
                web.Context.Load(folder, f => f.Exists);
                await web.Context.ExecuteQueryAsync();
                return folder.Exists;
            }
            catch (ServerException ex) when (ex.ServerErrorCode == -2147024894)
            {
                return false;
            }
        }
        private async Task<bool> IsFileExists(Web web, string fileUrl)
        {
            try
            {
                SP.File file = null;
                if (fileUrl.Contains("http"))
                    file = web.GetFileByUrl(fileUrl);
                else
                    file = web.GetFileByServerRelativeUrl(fileUrl);

                web.Context.Load(file, f => f.Exists);
                await web.Context.ExecuteQueryAsync();
                return file.Exists;
            }
            catch (ServerException ex) when (ex.ServerErrorCode == -2147024894)
            {
                return false;
            }
        }

    }
}
