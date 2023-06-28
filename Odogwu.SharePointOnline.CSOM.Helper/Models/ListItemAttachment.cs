using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class ListItemAttachment
    {
        public int ItemId { get; set; }
        public string FileName { get; set; }
        public string ServerRelativeUrl { get; set; }
        public string AbsoluteUrl { get; set; }
    }
}
