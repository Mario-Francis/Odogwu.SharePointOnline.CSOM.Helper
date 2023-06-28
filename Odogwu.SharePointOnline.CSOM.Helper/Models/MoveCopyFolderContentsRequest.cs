using System;
using System.Collections.Generic;
using System.Text;
using Odogwu.SharePointOnline.CSOM.Helper.Models.Enums;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class MoveCopyFolderContentsRequest
    {
        public string Library { get; set; }
        public string SourceFolder { get; set; }
        public string DestinationFolder { get; set; }
        public MoveCopyFolderContentOptions MoveCopyFolderContentOption { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
        public FolderContentTypes MoveCopyContentType { get; set; }
    }
}
