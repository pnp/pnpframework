using System;

namespace PnP.Framework.Migration.Lists.Planning
{
    internal static class ListFieldTypeCompatibility
    {
        public static bool IsCompatibleRuntimeType(string target, string source)
        {
            if (string.Equals(target, source, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            var targetFamily = Family(target);
            return targetFamily != null && string.Equals(targetFamily, Family(source), StringComparison.Ordinal);
        }

        private static string Family(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            switch (value.Trim().ToUpperInvariant())
            {
                case "TEXT":
                case "NOTE":
                case "CHOICE":
                    return "string";
                case "MULTICHOICE":
                    return "string-collection";
                case "INTEGER":
                case "COUNTER":
                case "NUMBER":
                case "CURRENCY":
                    return "number";
                case "BOOLEAN":
                    return "boolean";
                case "DATETIME":
                    return "datetime";
                case "GUID":
                    return "guid";
                case "URL":
                    return "url";
                case "USER":
                    return "user";
                case "USERMULTI":
                    return "user-collection";
                case "LOOKUP":
                    return "lookup";
                case "LOOKUPMULTI":
                    return "lookup-collection";
                case "TAXONOMYFIELDTYPE":
                    return "taxonomy";
                case "TAXONOMYFIELDTYPEMULTI":
                    return "taxonomy-collection";
                case "CALCULATED":
                case "COMPUTED":
                    return "computed";
                default:
                    return null;
            }
        }
    }
}
