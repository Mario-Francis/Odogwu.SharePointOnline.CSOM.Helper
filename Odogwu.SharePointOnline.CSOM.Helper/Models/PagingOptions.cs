using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class PagingOptions
    {
        public int StartIndex { get; set; } = 0;
        public int Length { get; set; } = 10;
        public bool SortByDateCreated { get; set; } = false;
        public string SortByDateCreatedDir { get; set; } = "ASC";
    }
}
