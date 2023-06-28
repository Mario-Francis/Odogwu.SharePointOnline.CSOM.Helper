using System;
using System.Collections.Generic;
using System.Text;
using Odogwu.SharePointOnline.CSOM.Helper.Models.Enums;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class MoveFileRequest
    {
        public string Library { get; set; }
        public string SourceFileUrl { get; set; }
        public string DestinationFolder { get; set; }
        public MoveCopyFileOptions MoveFileOption { get; set; }
        public bool? CreateDestinationFolderIfNotExist { get; set; }
    }
}
