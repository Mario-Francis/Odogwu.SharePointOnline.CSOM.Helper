using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class SPFile
    {
        public int Id { get; set; }
        public string Guid { get; set; }
        public string FileName { get; set; }
        public string ServerRelativeUrl { get; set; }
        public string AbsoluteUrl { get; set; }
        public byte[] File { get; set; }
        public string FileExtension { get; set; }
        public IEnumerable<KeyValuePair<string, object>> Properties { get; set; }
    }
}
