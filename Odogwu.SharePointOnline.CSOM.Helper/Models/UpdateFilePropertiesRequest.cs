using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class UpdateFilePropertiesRequest
    {
        public string Library { get; set; }
        public int Id { get; set; }
        public IEnumerable<KeyValuePair<string, object>> Properties { get; set; }

    }
}
