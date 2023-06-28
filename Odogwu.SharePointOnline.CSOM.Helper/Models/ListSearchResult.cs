using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class ListSearchResult
    {
        public IEnumerable<SPListItem> ListItems { get; set; }
        public int TotalResultCount { get; set; }
        public int TotalDisplayCount { get; set; }
    }
}
