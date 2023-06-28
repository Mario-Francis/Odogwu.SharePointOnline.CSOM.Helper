using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class AddItemRequest
    {
        public string ListName { get; set; }
        public IEnumerable<KeyValuePair<string, object>> FieldValues { get; set; }
    }
}
