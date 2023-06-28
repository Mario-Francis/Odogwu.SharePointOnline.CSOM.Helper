using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class SearchFilesByPropertiesRequest
    {
        
        public string Library { get; set; }
        public IEnumerable<KeyValuePair<string, object>> SearchParams  { get; set; }
        public PagingOptions Options { get; set; }
        public bool IncludeProperties { get; set; } = false;
        public string TargetFolder { get; set; }

        //public IEnumerable<KeyValuePair<string, string>> Sanitize()
        //{
        //    var paramList = new List<KeyValuePair<string, string>>();
        //    foreach(var p in SearchParams)
        //    {
        //        var key = p.Key.Replace("'", "").Replace("\"", "");
        //        var val = p.Value?.Replace("'", "").Replace("\"", "");
        //        paramList.Add(new KeyValuePair<string, string>(key, val));
        //    }
        //    return paramList;
        //}
    }
}
