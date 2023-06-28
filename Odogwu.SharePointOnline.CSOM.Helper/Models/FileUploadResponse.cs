using System.Collections.Generic;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class FileUploadResponse
    {
        public int Id { get; set; }
        public string Guid { get; set; }
        public string FileName { get; set; }
        public string ServerRelativeUrl { get; set; }
        public string AbsoluteUrl { get; set; }
    }
}
