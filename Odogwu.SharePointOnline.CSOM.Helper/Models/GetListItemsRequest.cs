using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class GetListItemsRequest
    {
        public string ListName { get; set; }
        public PagingOptions Options { get; set; }
    }
}
