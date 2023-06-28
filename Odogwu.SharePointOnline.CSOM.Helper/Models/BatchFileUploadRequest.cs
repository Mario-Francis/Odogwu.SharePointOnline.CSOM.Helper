using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class BatchFileUploadRequest
    {
        /// <summary>
        /// Document  library name
        /// </summary>
        public string Library { get; set; }

        /// <summary>
        /// Destination folder is the directory where the files will reside. The path should be relative to the library root folder. Leave blank/null for upload to root folder.
        /// </summary>
        public string DestinationFolder { get; set; }

        /// <summary>
        /// File upload items
        /// </summary>
        public IEnumerable<FileUploadItem> UploadItems { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
