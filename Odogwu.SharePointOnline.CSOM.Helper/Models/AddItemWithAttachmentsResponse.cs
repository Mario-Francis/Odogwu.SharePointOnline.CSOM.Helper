using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class AddItemWithAttachmentsResponse
    {
        public int Id { get; set; }
        public IEnumerable<ListItemAttachment> Attachments { get; set; }
    }
}
