using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public class Utilities
    {
        public static string GenerateCAMLAndTree(List<KeyValuePair<string, object>> conditions)
        {
            var total = conditions.Count;
            var opened = 0;
            var itemsLeft = total;
            var itemsAdded = 0;

            var xml = "";
            var count = ((int)total / 2) + 1;

            for (int i = 0; i <= count; i++)
            {
                // var added = 0;
                if (opened != 0)
                {
                    for (var j = opened; j > 0; j--)
                    {
                        xml += "</And>";
                    }
                }
                opened = 0;
                if (total == 1 && itemsLeft > 0)
                {
                    int index = itemsAdded;
                    //xml += $"<Eq><FieldRef Name='{conditions[index].Key}'/><Value Type='Text'>{conditions[index].Value}</Value></Eq>";
                    xml += GetEqTag(conditions[index].Key, conditions[index].Value);
                    itemsAdded += 1;
                    itemsLeft -= 1;
                }
                else if (itemsLeft > 1)
                {
                    if (itemsAdded == 4)
                    {
                        xml = "<And>" + xml + "</And>";
                        xml = "<And>" + xml;
                        opened += 1;
                    }
                    else if (itemsAdded > 4)
                    {
                        xml = "<And>" + xml;
                        opened += 1;
                    }
                    xml += "<And>";
                    opened += 1;
                    int index = itemsAdded;
                    //xml += $"<Eq><FieldRef Name='{conditions[index].Key}'/><Value Type='Text'>{conditions[index].Value}</Value></Eq>";
                    //xml += $"<Eq><FieldRef Name='{conditions[index + 1].Key}'/><Value Type='Text'>{conditions[index + 1].Value}</Value></Eq>";
                    xml += GetEqTag(conditions[index].Key, conditions[index].Value);
                    xml += GetEqTag(conditions[index + 1].Key, conditions[index + 1].Value);
                    itemsAdded += 2;
                    itemsLeft -= 2;

                    if (itemsAdded == 4 && total == 4)
                    {
                        xml = "<And>" + xml;
                        opened += 1;
                    }
                }
                else if (itemsLeft == 1)
                {
                    if (itemsAdded == 4)
                    {
                        xml = "<And>" + xml + "</And>";
                    }
                    xml = "<And>" + xml;
                    opened += 1;
                    int index = itemsAdded;
                    //xml += $"<Eq><FieldRef Name='{conditions[index].Key}'/><Value Type='Text'>{conditions[index].Value}</Value></Eq>";
                    xml += GetEqTag(conditions[index].Key, conditions[index].Value);
                    itemsAdded += 1;
                    itemsLeft -= 1;
                }
            }
            return xml;
        }

        private static string GetEqTag(string key, object val)
        {
            if (val == null)
            {
                return $"<Or><IsNull><FieldRef Name='{key}' /></IsNull><Eq><FieldRef Name='{key}'/><Value Type='{GetSPType(val)}'>{val}</Value></Eq></Or>";
            }
            else
            {
                return $"<Eq><FieldRef Name='{key}'/><Value Type='{GetSPType(val)}'>{val}</Value></Eq>";
            }
        }

        // private static string GetSPType(object val) => val.GetType() switch
        // {
        //     { } x when x == typeof(int) => "Integer",
        //     { } x when x == typeof(long) => "Integer",
        //     { } x when x == typeof(decimal) => "Currency",
        //     { } x when x == typeof(double) => "Currency",
        //     { } x when x == typeof(bool) => "Boolean",
        //     { } x when x == typeof(DateTime) => "DateTime",
        //     { } x when x == typeof(DateTimeOffset) => "DateTime",
        //     _ => "Text"
        // };

        private static string GetSPType(object val)
        {
            switch (val.GetType())
            {
                case var v when v == typeof(int):
                case var v2 when v2 ==  typeof(long):
                    return "Integer";
                case var v3 when v3 ==  typeof(decimal):
                case var v4 when v4 ==  typeof(double):
                    return "Currency";
                case var v5 when v5 ==  typeof(bool):
                    return "Boolean";
                case var v6 when v6 ==  typeof(DateTime):
                case var v7 when v7 ==  typeof(DateTimeOffset):
                    return "DateTime";
                default:
                    return "Text";
            }
        }

        public static string GetOrderByDateCreatedXml(string dir)
        {
            var isAscending = dir == "ASC";
            var val = isAscending ? "True" : "False";
            return $"<OrderBy><FieldRef Name=\"ID\" Ascending=\"{val}\" /></OrderBy>";
        }

        public static string GenerateUniqueId()
        {
            //var ticks = DateTime.Now.Ticks;
            var guid = Guid.NewGuid().ToString();
            return guid;
        }

        public static bool ValidateFieldNames(IEnumerable<string> listFields, IEnumerable<string> propFields, out string invalidField)
        {
            var res = true;
            invalidField = "";
            foreach (var pf in propFields)
            {
                if (!listFields.Any(l => l == pf))
                {
                    res = false;
                    invalidField = pf;
                    break;
                }
            }
            return res;
        }

    }
}
