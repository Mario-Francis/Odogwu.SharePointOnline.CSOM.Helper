using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class SearchListByFieldValuesRequest
    {
        public string ListName { get; set; }
        public IEnumerable<KeyValuePair<string, object>> SearchParams { get; set; }
        public PagingOptions Options { get; set; }
    }
}
