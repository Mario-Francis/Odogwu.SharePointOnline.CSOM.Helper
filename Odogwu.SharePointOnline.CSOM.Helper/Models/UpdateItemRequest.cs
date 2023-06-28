using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class UpdateItemRequest
    {
        public string ListName { get; set; }
        public int Id { get; set; }
        public IEnumerable<KeyValuePair<string, object>> FieldValues { get; set; }
    }
}
