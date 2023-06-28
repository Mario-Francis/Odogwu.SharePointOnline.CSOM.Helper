using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class CopyFileRequest
    {
        public string Library { get; set; }
        public string SourceFileUrl { get; set; }
        public string DestinationFolder { get; set; }
        public string NewFileName { get; set; }
        public MoveCopyFileOptions CopyFileOption { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
