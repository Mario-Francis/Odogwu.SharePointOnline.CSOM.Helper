using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class CopyFolderRequest
    {
        public string Library { get; set; }
        public string SourceFolderUrl { get; set; }
        public string DestinationFolder { get; set; }
        public string NewFolderName { get; set; }
        public MoveCopyFolderOptions CopyFolderOption { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
