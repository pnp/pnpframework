using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.Utilities
{
    public static class ListItemUtilities
    {
        private readonly static string[] WriteableReadOnlyFields = new string[] { "publishingpagelayout", "contenttypeid", "description" };

        public static FieldUpdateValue ParseFieldValue(Web web, string fieldValue, Field dataField)
        {
            switch (dataField.FieldTypeKind)
            {
                case FieldType.Geolocation:
                    // FieldGeolocationValue - Expected format: Altitude,Latitude,Longitude,Measure
                    var geolocationArray = fieldValue.Split(',');
                    if (geolocationArray.Length == 4)
                    {
                        var geolocationValue = new FieldGeolocationValue
                        {
                            Altitude = Double.Parse(geolocationArray[0]),
                            Latitude = Double.Parse(geolocationArray[1]),
                            Longitude = Double.Parse(geolocationArray[2]),
                            Measure = Double.Parse(geolocationArray[3]),
                        };
                        return new FieldUpdateValue(dataField.InternalName, geolocationValue);
                    }
                    else
                    {
                        return new FieldUpdateValue(dataField.InternalName, fieldValue);
                    }
                case FieldType.Lookup:
                    if (dataField.TypeAsString == "LookupMulti" && TryDeserializeAsJson(fieldValue, out FieldLookupValue[] lookupValues2))
                    {
                        // Backward compatibility, when format was stored as json
                        return new FieldUpdateValue(dataField.InternalName, lookupValues2);
                    }
                    // FieldLookupValue - Expected format: LookupID or LookupID,LookupID,LookupID...
                    else if (fieldValue.Contains(","))
                    {
                        var lookupValues = new List<FieldLookupValue>();
                        fieldValue.Split(',').All(value =>
                        {
                            lookupValues.Add(new FieldLookupValue
                            {
                                LookupId = int.Parse(value),
                            });
                            return true;
                        });
                        return new FieldUpdateValue(dataField.InternalName, lookupValues.ToArray());
                    }
                    else
                    {
                        var lookupValue = new FieldLookupValue
                        {
                            LookupId = int.Parse(fieldValue),
                        };
                        return new FieldUpdateValue(dataField.InternalName, lookupValue);
                    }
                case FieldType.URL:
                    // FieldUrlValue - Expected format: URL,Description
                    var urlArray = fieldValue.Split(new Char[] { ',' }, 2);
                    var linkValue = new FieldUrlValue();
                    if (urlArray.Length == 2)
                    {
                        linkValue.Url = urlArray[0];
                        linkValue.Description = urlArray[1];
                    }
                    else
                    {
                        linkValue.Url = urlArray[0];
                        linkValue.Description = urlArray[0];
                    }
                    return new FieldUpdateValue(dataField.InternalName, linkValue);

                case FieldType.User:
                    // FieldUserValue - Expected format: loginName or loginName,loginName,loginName...
                    if (fieldValue.Contains(","))
                    {
                        var userValues = new List<FieldUserValue>();
                        fieldValue.Split(',').All(value =>
                        {
                            var user = web.EnsureUser(value);
                            web.Context.Load(user);
                            web.Context.ExecuteQueryRetry();
                            if (user != null)
                            {
                                userValues.Add(new FieldUserValue
                                {
                                    LookupId = user.Id,
                                });
                            }
                            return true;
                        });
                        return new FieldUpdateValue(dataField.InternalName, userValues.ToArray());
                    }
                    else
                    {
                        var user = web.EnsureUser(fieldValue);
                        web.Context.Load(user);
                        web.Context.ExecuteQueryRetry();
                        if (user != null)
                        {
                            var userValue = new FieldUserValue
                            {
                                LookupId = user.Id,
                            };
                            return new FieldUpdateValue(dataField.InternalName, userValue);
                        }
                        else
                        {
                            return new FieldUpdateValue(dataField.InternalName, fieldValue);
                        }
                    }
                case FieldType.DateTime:
                    var dateTime = DateTime.MinValue;
                    if (DateTime.TryParse(fieldValue, out dateTime))
                    {
                        return new FieldUpdateValue(dataField.InternalName, dateTime);
                    }
                    break;

                case FieldType.MultiChoice:
                    if (TryDeserializeAsJson(fieldValue, out string[] choices))
                    {
                        // Backward compatibility: format is json
                        return new FieldUpdateValue(dataField.InternalName, choices);
                    }
                    else
                    {
                        // expected format: Choice1;#Choice2;#Choice3,
                        return new FieldUpdateValue(dataField.InternalName, fieldValue.Split(";#"));
                    }
                case FieldType.Invalid:
                    switch (dataField.TypeAsString)
                    {
                        case "TaxonomyFieldType":
                            {
                                if (fieldValue != null && TryDeserializeAsJson(fieldValue, out TaxonomyFieldValue taxVal))
                                {
                                    // Backward compatibility, when format was stored as json
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        new List<string> { $"-1;#{taxVal.Label}|{taxVal.TermGuid}" },
                                        dataField.TypeAsString
                                        );
                                }
                                else if (fieldValue != null)
                                {
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        new List<string> { $"-1;#{fieldValue}" },
                                        dataField.TypeAsString
                                        );
                                }
                                break;
                            }
                        // Single value field - Expected format: term label|term GUID
                        case "TaxonomyFieldTypeMulti":
                            {
                                if (fieldValue != null && TryDeserializeAsJson(fieldValue, out TaxonomyFieldValue[] taxValues))
                                {
                                    // Backward compatibility, when format was stored as json
                                    return new FieldUpdateValue(
                                        dataField.InternalName,
                                        taxValues.Select(taxVal => $"-1;#{taxVal.Label}|{taxVal.TermGuid}"),
                                        dataField.TypeAsString
                                        );

                                }
                                else if (fieldValue != null)
                                {
                                    // Multi value field - Expected format: term label|term GUID;term label|term GUID;term label|term GUID;...
                                    var termStrings = new List<string>();

                                    var termsArray = fieldValue.Split(new char[] { ';' });
                                    foreach (var term in termsArray)
                                    {
                                        termStrings.Add($"-1;#{term}");
                                    }
                                    return new FieldUpdateValue(dataField.InternalName, termStrings, dataField.TypeAsString);
                                }
                                break;
                            }
                    }
                    break;
            }

            // Default to set direct value
            return new FieldUpdateValue(dataField.InternalName, fieldValue, dataField.TypeAsString);
        }

        [Obsolete("Use UpdateListItem(ListItem item, TokenParser parser, IDictionary<string, string> valuesToSet, ListItemUpdateType updateType) instead")]
        public static void UpdateListItem(
            Web web,
            ListItem listitem,
            TokenParser parser,
            FieldCollection fields,
            IDictionary<string, string> fieldValues
            )
        {
            var updateValues = new List<FieldUpdateValue>();

            foreach (var dataValue in fieldValues)
            {
                Field dataField = null;

                if (parser != null)
                {
                    dataField = fields.FirstOrDefault(f => f.InternalName == parser.ParseString(dataValue.Key));
                }
                else
                {
                    dataField = fields.FirstOrDefault(f => f.InternalName == dataValue.Key);
                }

                if (dataField == null)
                {
                    // TODO: log Warning
                    continue;
                }

                // Changed by PaoloPia because there are fields like PublishingPageLayout
                // which are marked as read-only, but have to be overwritten while uploading
                // a publishing page file and which in reality can still be written
                if (
                    dataField.ReadOnlyField
                    && !WriteableReadOnlyFields.Contains(dataField.InternalName.ToLower()))
                {
                    // skip read only fields
                    continue;
                }

                if (dataValue.Value == null)
                {
                    updateValues.Add(new FieldUpdateValue(dataValue.Key, null, dataField.TypeAsString));
                }
                else
                {
                    var fieldValue = parser.ParseString(dataValue.Value);

                    updateValues.Add(
                        ParseFieldValue(web, fieldValue, dataField)
                        );
                }
            }

            UpdateListItem(web, listitem, fields, updateValues);
        }

        public enum ListItemUpdateType
        {
            Update,
            SystemUpdate,
            UpdateOverwriteVersion
        }

        public static void UpdateListItem(ListItem item, TokenParser parser, IDictionary<string, string> valuesToSet, ListItemUpdateType updateType, bool SkipExecuteQuery = false)
        {
            var itemValues = new List<FieldUpdateValue>();

            var context = item.Context as ClientContext;
            var list = item.ParentList;
            context.Web.EnsureProperty(w => w.Url);

            bool isDocLib = list.EnsureProperty(l => l.BaseType) == BaseType.DocumentLibrary;
            bool isPagesLib = list.EnsureProperty(l => l.RootFolder).Name.Equals("SitePages", StringComparison.InvariantCultureIgnoreCase);

            var clonedContext = context.Clone(context.Web.Url);
            var web = clonedContext.Web;

            var fields =
                     context.LoadQuery(list.Fields.Include(f => f.InternalName, f => f.Title,
                         f => f.TypeAsString));
            context.ExecuteQueryRetry();

            Regex fileUniqueIdToken = new Regex("(?<token>[{]{1,2}(?:fileuniqueid:fileuniqueidencoded:)|[^}]*[}]{1,2})(?:[^{]*)", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);

            foreach (var key in valuesToSet.Keys)
            {
                var field = fields.FirstOrDefault(f => f.InternalName == key || f.Title == key);

                if (field != null)
                {
                    if (field.InternalName.Equals("ID", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // Ignor ID column. ID column cannot be updated (Exception: This field cannot be updated.)
                        // Sometimes the ID column is used as KeyColumn for DataRows.
                        continue;
                    }

                    var value = parser.ParseString(valuesToSet[key]);

                    if (!string.IsNullOrWhiteSpace(value) && fileUniqueIdToken.IsMatch(value))
                    {
                        //unresovled fileuniqueid or fileuniqueidencoded Token -try to reslove and add to parser..
                        foreach (Match m in fileUniqueIdToken.Matches(value))
                        {
                            if (m.Groups["token"].Success && !string.IsNullOrWhiteSpace(m.Groups["token"].Value))
                            {
                                string replacementVal = m.Groups["token"].Value;
                                var tokenParts = replacementVal.Trim(new char[] { '{', '}' }).Split(':');

                                // note 1: the RegEx can match content that contains no valid tokens; do some basic sanity checks
                                // note 2: checking for empty tokenParts[1] because this would return an actual folder below when calling GetFolderByServerRelativePath, which would be an error
                                if (tokenParts.Length < 2 || string.IsNullOrWhiteSpace(tokenParts[1]))
                                {
                                    continue;
                                }

                                Guid UniqueId = Guid.Empty;
                                try
                                {
                                    var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(tokenParts[1]));
                                    file.EnsureProperties(f => f.UniqueId);
                                    UniqueId = file.UniqueId;
                                }
                                catch { }

                                if (UniqueId == Guid.Empty)
                                {
                                    try
                                    {
                                        var folder = web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(tokenParts[1]));
                                        folder.EnsureProperties(f => f.UniqueId);
                                        UniqueId = folder.UniqueId;
                                    }
                                    catch { }
                                }

                                if (UniqueId != Guid.Empty)
                                {
                                    if (tokenParts[0].Equals("fileuniqueidencoded", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        string replTest = UniqueId.ToString().Replace("-", "%2D");
                                        value = Regex.Replace(value, $"{{{tokenParts[0]}:{tokenParts[1]}}}", replTest);
                                    }
                                    else
                                    {
                                        string replTest = UniqueId.ToString();
                                        value = Regex.Replace(value, $"{{{tokenParts[0]}:{tokenParts[1]}}}", replTest);
                                    }
                                }
                            }
                        }
                    }

                    switch (field.TypeAsString)
                    {
                        case "User":
                        case "UserMulti":
                            {
                                List<FieldUserValue> userValues = new List<FieldUserValue>();

                                if (value == null) goto default;
                                if (value is string && string.IsNullOrWhiteSpace(value + "")) goto default;

                                if (value.Contains(","))
                                {
                                    var valueArray = value.Split(new char[] { ',' });
                                    foreach (var arrayItem in valueArray)
                                    {
                                        if (!int.TryParse(arrayItem.Trim().ToString(), out int userId))
                                        {
                                            var user = web.EnsureUser(arrayItem.Trim());
                                            clonedContext.Load(user);
                                            clonedContext.ExecuteQueryRetry();
                                            userValues.Add(new FieldUserValue() { LookupId = user.Id });

                                        }
                                        else
                                        {
                                            userValues.Add(new FieldUserValue() { LookupId = userId });
                                        }
                                    }
                                    itemValues.Add(new FieldUpdateValue(key, userValues.ToArray(), null));
                                }
                                else
                                {
                                    if (!int.TryParse(value.Trim(), out int userId))
                                    {
                                        var user = web.EnsureUser(value.Trim());
                                        clonedContext.Load(user);
                                        clonedContext.ExecuteQueryRetry();
                                        itemValues.Add(new FieldUpdateValue(key, new FieldUserValue() { LookupId = user.Id }));
                                    }
                                    else
                                    {
                                        itemValues.Add(new FieldUpdateValue(key, new FieldUserValue() { LookupId = userId }));
                                    }
                                }
                                break;
                            }
                        case "MultiChoice":
                            {
                                if (value != null)
                                {
                                    var array = value.Split(";#");
                                    itemValues.Add(new FieldUpdateValue(key, array));
                                }
                                break;
                            }
                        case "TaxonomyFieldType":
                        case "TaxonomyFieldTypeMulti":
                            {
                                if (value != null && (value.Contains(",") || value.Contains(";")))
                                {
                                    var taxSession = clonedContext.Site.GetTaxonomySession();
                                    var terms = new List<KeyValuePair<Guid, string>>();
                                    foreach (var arrayItem in value.Split(new char[] { ',', ';' }))
                                    {
                                        TaxonomyItem taxonomyItem;
                                        if (!Guid.TryParse(arrayItem.Trim(), out Guid termGuid))
                                        {
                                            // Assume it's a TermPath
                                            taxonomyItem = clonedContext.Site.GetTaxonomyItemByPath(arrayItem.Trim());
                                        }
                                        else
                                        {
                                            taxonomyItem = taxSession.GetTerm(termGuid);
                                            clonedContext.Load(taxonomyItem);
                                            clonedContext.ExecuteQueryRetry();
                                        }
                                        if (taxonomyItem != null)
                                        {
                                            terms.Add(new KeyValuePair<Guid, string>(taxonomyItem.Id, taxonomyItem.Name));
                                        }
                                    }

                                    TaxonomyField taxField = context.CastTo<TaxonomyField>(field);
                                    taxField.EnsureProperty(tf => tf.AllowMultipleValues);
                                    if (taxField.AllowMultipleValues)
                                    {
                                        var termValuesString = String.Empty;
                                        foreach (var term in terms)
                                        {
                                            termValuesString += "-1;#" + term.Value + "|" + term.Key.ToString("D") + ";#";
                                        }

                                        termValuesString = termValuesString.Substring(0, termValuesString.Length - 2);

                                        var newTaxFieldValue = new TaxonomyFieldValueCollection(context, termValuesString, taxField);
                                        itemValues.Add(new FieldUpdateValue(key, newTaxFieldValue, field.TypeAsString));
                                    }
                                }
                                else
                                {
                                    Guid termGuid = Guid.Empty;

                                    var taxSession = clonedContext.Site.GetTaxonomySession();
                                    TaxonomyItem taxonomyItem = null;
                                    if (!string.IsNullOrEmpty(value) && !Guid.TryParse(value.Trim(), out termGuid))
                                    {
                                        // Assume it's a TermPath
                                        taxonomyItem = clonedContext.Site.GetTaxonomyItemByPath(value.Trim());
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(value))
                                        {
                                            taxonomyItem = taxSession.GetTerm(termGuid);
                                            clonedContext.Load(taxonomyItem);
                                            clonedContext.ExecuteQueryRetry();
                                        }
                                    }

                                    TaxonomyField taxField = context.CastTo<TaxonomyField>(field);
                                    TaxonomyFieldValue taxValue = new TaxonomyFieldValue();
                                    if (taxonomyItem != null)
                                    {
                                        taxValue.TermGuid = taxonomyItem.Id.ToString();
                                        taxValue.Label = taxonomyItem.Name;
                                        itemValues.Add(new FieldUpdateValue(key, taxValue, field.TypeAsString));
                                    }
                                    else
                                    {
                                        taxField.ValidateSetValue(item, null);
                                    }
                                }
                                break;
                            }
                        case "Lookup":
                        case "LookupMulti":
                            {
                                if (value == null) goto default;

                                var newVals = value.Split(',', ';')
                                    .Select(v => new FieldLookupValue { LookupId = int.Parse(v.Trim()) })
                                    .ToArray();

                                var lookupField = context.CastTo<FieldLookup>(field);
                                lookupField.EnsureProperty(lf => lf.AllowMultipleValues);
                                if (!lookupField.AllowMultipleValues && newVals.Length > 1)
                                {
                                    throw new Exception("Field " + field.InternalName + " does not support multiple values");
                                }
                                itemValues.Add(new FieldUpdateValue(key, newVals));
                                break;
                            }
                        case "DateTime":
                            {
                                if (value == null) goto default;
                                if (DateTime.TryParse(value, out DateTime dateTimeValue))
                                {
                                    itemValues.Add(new FieldUpdateValue(key, dateTimeValue));
                                }
                                break;
                            }
                        case "URL":
                            {

                                if (value == null) goto default;
                                if (value.Contains(",") || value.Contains(";"))
                                {
                                    var urlValueArray = value.Split(new char[] { ',', ';' });
                                    if (urlValueArray.Length == 2)
                                    {
                                        var urlValue = new FieldUrlValue
                                        {
                                            Url = value.Split(new char[] { ',', ';' })[0],
                                            Description = value.Split(new char[] { ',', ';' })[1]
                                        };
                                        itemValues.Add(new FieldUpdateValue(key, urlValue));
                                    }
                                    else
                                    {
                                        itemValues.Add(new FieldUpdateValue(key, value));
                                    }
                                }
                                else
                                {
                                    var urlValue = new FieldUrlValue
                                    {
                                        Url = value,
                                        Description = value
                                    };
                                    itemValues.Add(new FieldUpdateValue(key, urlValue));
                                }

                                break;
                            }
                        case "Geolocation":
                            {
                                if (value == null) goto default;
                                // FieldGeolocationValue - Expected format: Altitude,Latitude,Longitude,Measure
                                var geolocationArray = value.Split(',');
                                if (geolocationArray.Length == 4)
                                {
                                    var geolocationValue = new FieldGeolocationValue
                                    {
                                        Altitude = Double.Parse(geolocationArray[0]),
                                        Latitude = Double.Parse(geolocationArray[1]),
                                        Longitude = Double.Parse(geolocationArray[2]),
                                        Measure = Double.Parse(geolocationArray[3]),
                                    };

                                    itemValues.Add(new FieldUpdateValue(key, geolocationValue));
                                }
                                else
                                {
                                    itemValues.Add(new FieldUpdateValue(key, value));
                                }
                               
                                break;
                            }
                        default:
                            {
                                itemValues.Add(new FieldUpdateValue(key, value));
                                break;
                            }
                    }
                }
            }

            if (isDocLib)
            {
                // check if we have both editor and author in the item.
                var setAuthor = itemValues.FirstOrDefault(v => v.Key.Equals("author", StringComparison.InvariantCultureIgnoreCase)) != null;
                var setEditor = itemValues.FirstOrDefault(v => v.Key.Equals("editor", StringComparison.InvariantCultureIgnoreCase)) != null;
                if ((!setAuthor || !setEditor) && (setAuthor || setEditor))
                {
                    if (!setAuthor)
                    {
                        var currentAuthor = item["Author"];
                        // the null check catches the case where somebody tries to add new list items to a doc lib and the server says No
                        if (currentAuthor != null)
                        {
                            // We only have the editor field, set the author to the old value
                            itemValues.Add(new FieldUpdateValue("Author", currentAuthor));
                        }
                    }
                    if (!setEditor)
                    {
                        var currentEditor = item["Editor"];
                        // the null check catches the case where somebody tries to add new list items to a doc lib and the server says No
                        if (currentEditor != null)
                        {
                            // We opnly have the author field, set the editor to the old value
                            itemValues.Add(new FieldUpdateValue("Editor", currentEditor));
                        }
                    }
                }
            }
            foreach (var itemValue in itemValues)
            {
                if (string.IsNullOrEmpty(itemValue.FieldTypeString))
                {
                    item[itemValue.Key] = itemValue.Value;
                }
                else
                {
                    switch (itemValue.FieldTypeString)
                    {
                        case "TaxonomyFieldTypeMulti":
                            {
                                var field = fields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                                var taxField = context.CastTo<TaxonomyField>(field);
                                if (itemValue.Value is TaxonomyFieldValueCollection)
                                {
                                    taxField.SetFieldValueByValueCollection(item, itemValue.Value as TaxonomyFieldValueCollection);
                                }
                                else
                                {
                                    taxField.SetFieldValueByValue(item, itemValue.Value as TaxonomyFieldValue);
                                }

                                break;
                            }
                        case "TaxonomyFieldType":
                            {
                                var field = fields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                                var taxField = context.CastTo<TaxonomyField>(field);
                                taxField.SetFieldValueByValue(item, itemValue.Value as TaxonomyFieldValue);
                                break;
                            }
                    }
                }
            }
            switch (updateType)
            {
                case ListItemUpdateType.Update:
                    {
                        item.Update();
                        break;
                    }
                case ListItemUpdateType.SystemUpdate:
                    {
                        item.SystemUpdate();
                        break;
                    }
                case ListItemUpdateType.UpdateOverwriteVersion:
                    {
                        var itemIsModernClientSidePage = isPagesLib && item["File_x0020_Type"]?.ToString() == "aspx";
                        if (itemIsModernClientSidePage)
                        {
                            // when updating fields of modern client side pages UpdateOverwriteVersion throws this error: "Additions to this Web site have been blocked."
                            // so use SystemUpdate instead
                            item.SystemUpdate();
                        }
                        else
                        {
                            item.UpdateOverwriteVersion();
                        }
                        break;
                    }
            }

            if (!SkipExecuteQuery)
                context.ExecuteQueryRetry();
        }

        [Obsolete("Use UpdateListItem(ListItem item, TokenParser parser, IDictionary<string, string> valuesToSet, ListItemUpdateType updateType) instead")]
        public static void UpdateListItem(
            Web web,
            ListItem listItem,
            FieldCollection listFields,
            IEnumerable<FieldUpdateValue> updateValues
            )
        {
            if (web == null) throw new ArgumentNullException(nameof(web));
            if (listFields == null) throw new ArgumentNullException(nameof(listFields));
            if (listItem == null) throw new ArgumentNullException(nameof(listItem));

            if (updateValues == null || !updateValues.Any()) return;

            foreach (var itemValue in updateValues.Where(u => u.FieldTypeString != "TaxonomyFieldTypeMulti" && u.FieldTypeString != "TaxonomyFieldType"))
            {
                // Special case for ContentType field
                if (itemValue.Key == "ContentType")
                {
                    var targetCT = listItem.ParentList.GetContentTypeByName((string)itemValue.Value);
                    web.Context.ExecuteQueryRetry();

                    if (targetCT != null)
                    {
                        listItem["ContentTypeId"] = targetCT.StringId;
                    }
                    else
                    {
                        Log.Error(Constants.LOGGING_SOURCE, "Content Type {0} does not exist in target list!", (string)itemValue.Value);
                    }
                }
                else
                {
                    listItem[itemValue.Key] = itemValue.Value;
                }
            }
            listItem.Update();
            web.Context.Load(listItem);
            web.Context.ExecuteQueryRetry();
            var itemId = listItem.Id;
            foreach (var itemValue in updateValues.Where(u => u.FieldTypeString == "TaxonomyFieldTypeMulti" || u.FieldTypeString == "TaxonomyFieldType"))
            {
                switch (itemValue.FieldTypeString)
                {
                    case "TaxonomyFieldTypeMulti":
                        {
                            var field = listFields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                            var taxField = web.Context.CastTo<TaxonomyField>(field);
                            if (itemValue.Value != null)
                            {
                                var valueCollection = new TaxonomyFieldValueCollection(web.Context, string.Join(";#", itemValue.Value as IEnumerable<string>), taxField);
                                taxField.SetFieldValueByValueCollection(listItem, valueCollection);
                            }
                            else
                            {
                                var valueCollection = new TaxonomyFieldValueCollection(web.Context, null, taxField);
                                taxField.SetFieldValueByValueCollection(listItem, valueCollection);
                            }
                            listItem.Update();
                            web.Context.Load(listItem);
                            web.Context.ExecuteQueryRetry();
                            break;
                        }
                    case "TaxonomyFieldType":
                        {
                            var field = listFields.FirstOrDefault(f => f.InternalName == itemValue.Key || f.Title == itemValue.Key);
                            var taxField = web.Context.CastTo<TaxonomyField>(field);
                            taxField.EnsureProperty(f => f.TextField);
                            var taxValue = new TaxonomyFieldValue();
                            if (itemValue.Value != null)
                            {
                                var termString = ((List<string>)itemValue.Value)[0];
                                taxValue.Label = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[0];
                                taxValue.TermGuid = termString.Split(new string[] { ";#" }, StringSplitOptions.None)[1].Split(new char[] { '|' })[1];
                                taxValue.WssId = -1;
                                taxField.SetFieldValueByValue(listItem, taxValue);
                            }
                            else
                            {
                                taxValue.Label = string.Empty;
                                taxValue.TermGuid = "11111111-1111-1111-1111-111111111111";
                                taxValue.WssId = -1;
                                Field hiddenField = listFields.GetById(taxField.TextField);
                                listItem.Context.Load(hiddenField, tf => tf.InternalName);
                                listItem.Context.ExecuteQueryRetry();
                                taxField.SetFieldValueByValue(listItem, taxValue); // this order of updates is important.
                                listItem[hiddenField.InternalName] = string.Empty; // this order of updates is important.
                            }
                            listItem.Update();
                            web.Context.Load(listItem);
                            web.Context.ExecuteQueryRetry();
                            break;
                        }
                }
            }
        }

        /// <summary>
        /// This method is present to preserve backward compatibility with old file property format
        /// </summary>
        /// <typeparam name="T">Expected type</typeparam>
        /// <param name="jsonValue">json value</param>
        /// <param name="result">The result object, if success</param>
        /// <returns>Returns <c>true</c> if the value was sucessfully deserialized from the json string. Otherwise <c>false</c></returns>
        private static bool TryDeserializeAsJson<T>(string jsonValue, out T result)
        {
            try
            {
                result = JsonUtility.Deserialize<T>(jsonValue);
                return true;
            }
            catch (Newtonsoft.Json.JsonException)
            {
                result = default(T);
                return false;
            }
            // Other exception are not to be catched
        }
    }
}
