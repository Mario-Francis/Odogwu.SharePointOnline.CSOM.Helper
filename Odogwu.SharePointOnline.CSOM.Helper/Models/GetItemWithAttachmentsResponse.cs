using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class GetItemWithAttachmentsResponse
    {
        public SPListItem ListItem { get; set; }
        public IEnumerable<ListItemAttachment> Attachments { get; set; }
    }
}
