using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class AddItemWithAttachmentsRequest
    {
        public string ListName { get; set; }
        public IEnumerable<KeyValuePair<string, string>> ItemFieldValues { get; set; }
        public IEnumerable<AttachmentUploadItem> UploadItems { get; set; }
    }
}
