using System.Collections.Generic;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class FileUploadRequest
    {
        /// <summary>
        /// Document  library name
        /// </summary>
        public string Library { get; set; }
       
        /// <summary>
        /// Destination folder is the directory where the file will reside. The path should be relative to the library root folder. Leave blank/null for upload to root folder.
        /// </summary>
        public string DestinationFolder { get; set; }
       
        /// <summary>
        /// File upload item
        /// </summary>
        public FileUploadItem UploadItem { get; set; }

        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
