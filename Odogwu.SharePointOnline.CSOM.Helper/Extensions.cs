using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public static class Extensions
    {
        public static string GetSiteServerRelativeUrl(this ClientRuntimeContext context)
        {
            var spDomain = "sharepoint.com";
            var url = context.Url.Substring(context.Url.IndexOf(spDomain) + spDomain.Length);
            return url;
        }

        public static string GetServerBaseUrl(this ClientRuntimeContext context)
        {
            var url = context.Url.Replace(context.GetSiteServerRelativeUrl(), "");
            return url;
        }

        public static IEnumerable<KeyValuePair<string, object>> Sanitize(this IEnumerable<KeyValuePair<string, object>> param)
        {
            var paramList = new List<KeyValuePair<string, object>>();
            foreach (var p in param)
            {
                var key = p.Key.Replace("'", "").Replace("\"", "");
                object val = p.Value;
                if (p.Value != null && p.Value.GetType() == typeof(string))
                {
                    val = p.Value.ToString().Replace("'", "").Replace("\"", "");
                }
                paramList.Add(new KeyValuePair<string, object>(key, val));
            }
            return paramList;
        }

        // public static string ResolveTargetFolder(this ClientRuntimeContext context, string library, string targetFolderUrl) => targetFolderUrl?.Trim() switch
        // {
        //     null => null,
        //     string s when s.Length > 0 && s.StartsWith("/") && s.Trim('/').Length > 0 => $"{context.GetSiteServerRelativeUrl()}/{targetFolderUrl?.Trim('/')}",
        //     _ => $"{context.GetSiteServerRelativeUrl()}/{library}/{targetFolderUrl?.Trim('/')}"
        // };

        public static string ResolveTargetFolder(this ClientRuntimeContext context, string library, string targetFolderUrl)
        {
            switch (targetFolderUrl?.Trim())
            {
                case null:
                    return null;
                case string s when s.Length > 0 && s.StartsWith("/") && s.Trim('/').Length > 0:
                    return $"{context.GetSiteServerRelativeUrl()}/{targetFolderUrl?.Trim('/')}";
                default:
                    return $"{context.GetSiteServerRelativeUrl()}/{library}/{targetFolderUrl?.Trim('/')}";
            }
        }
    }
}
