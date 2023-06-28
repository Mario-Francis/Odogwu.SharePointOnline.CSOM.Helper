using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper.Models
{
    public class AttachmentUploadItem
    {
        /// <summary>
        /// Binary file
        /// </summary>
        public byte[] File { get; set; }
        /// <summary>
        /// File name
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// File extension
        /// </summary>
        public string FileExtension { get; set; }
        public string GetFileNameWithExtension()
        {
            var ext = FileExtension.StartsWith(".") ? FileExtension : "." + FileExtension;
            return FileName.Replace(" ", "_") + ext;
        }

        public string GetUniqueFileNameWithExtension()
        {
            var ext = FileExtension.StartsWith(".") ? FileExtension : "." + FileExtension;
            return FileName.Replace(" ", "_").Replace(".", "_") + "-" + Utilities.GenerateUniqueId() + ext;
        }
    }
}
