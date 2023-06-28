using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class MoveFolderRequest
    {
        public string Library { get; set; }
        public string SourceFolderUrl { get; set; }
        public string DestinationFolder { get; set; }
        public MoveCopyFolderOptions MoveFolderOption { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
