using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class BatchAddItemsRequest
    {
        public string ListName { get; set; }
        public IEnumerable<IEnumerable<KeyValuePair<string, object>>> Items { get; set; }

    }
}
