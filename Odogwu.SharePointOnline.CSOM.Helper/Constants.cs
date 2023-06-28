using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public enum MoveCopyFileOptions
    {
       RenameDuplicate,
       ReportDuplicate,
       OverwriteDuplicate
    }

    public enum MoveCopyFolderOptions
    {
        RenameDuplicate,
        ReportDuplicate
    }
    public enum MoveCopyFolderContentOptions
    {
        RenameDuplicate,
        ReportDuplicate
    }
    public enum FolderContentTypes
    {
        All,
        FilesOnly,
        FoldersOnly
    }

    class Constants
    {
    }
}
