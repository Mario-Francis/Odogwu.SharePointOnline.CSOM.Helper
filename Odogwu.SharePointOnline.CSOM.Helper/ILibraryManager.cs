using Odogwu.SharePointOnline.CSOM.Helper.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public interface ILibraryManager
    {
        string SiteUrl { get; set; }
        Task<FileUploadResponse> UploadFile(FileUploadRequest request);
        Task<IEnumerable<FileUploadResponse>> BatchUploadFile(BatchFileUploadRequest request, int maxUploadItem = 10);
        Task<SPFile> GetFileById(int id, string library);

        Task<SPFile> GetFileByUniqueId(string uniqueId, string library);

        Task<SPFile> GetFileByUrl(string fileUrl, string library);

        Task<FileSearchResult> SearchFilesByProperties(SearchFilesByPropertiesRequest request, int maxResultLength = 10, int maxItemLoad = 500000);

        Task<FileSearchResult> SearchFilesByDateRange(SearchFilesByDateRangeRequest request, int maxResultLength = 10, int maxItemLoad = 500000);

        Task<SPFile> UpdateFileProperties(UpdateFilePropertiesRequest request);

        Task DeleteFileById(int id, string library);
        Task DeleteFileByUrl(string fileUrl);
    }
}
