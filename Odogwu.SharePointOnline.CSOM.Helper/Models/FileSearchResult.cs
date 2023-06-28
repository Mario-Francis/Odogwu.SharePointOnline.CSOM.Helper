using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class FileSearchResult
    {
        public IEnumerable<SPFile> Files { get; set; }
        public int TotalResultCount { get; set; }
        public int TotalDisplayCount { get; set; }
    }
}
