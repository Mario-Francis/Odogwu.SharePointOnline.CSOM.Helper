using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class SearchFilesByDateRangeRequest
    {
        public string Library { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public PagingOptions Options { get; set; }
        public bool IncludeProperties { get; set; } = false;
        public string TargetFolder { get; set; }

    }
}
