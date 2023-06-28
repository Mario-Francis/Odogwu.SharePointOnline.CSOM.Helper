using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class SearchListByDateRangeRequest
    {
        public string ListName { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public PagingOptions Options { get; set; }
    }
}
