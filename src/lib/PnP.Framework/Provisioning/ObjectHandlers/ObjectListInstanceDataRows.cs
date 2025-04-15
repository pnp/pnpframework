﻿using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Extensions;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstanceDataRows : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "List instances Data Rows"; }
        }

        public override string InternalName => "ListInstanceDataRows";
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (!template.Lists.Any()) return parser;

                web.EnsureProperties(w => w.ServerRelativeUrl);

                web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                web.Context.ExecuteQueryRetry();

                #region DataRows

                foreach (var listInstance in template.Lists)
                {
                    if (listInstance.DataRows != null && listInstance.DataRows.Any())
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Processing_data_rows_for__0_, listInstance.Title);
                        // Retrieve the target list
                        var list = web.GetListByUrl(parser.ParseString(listInstance.Url));
                        web.Context.Load(list);

                        // Retrieve the fields' types from the list
                        Microsoft.SharePoint.Client.FieldCollection fields = list.Fields;
                        web.Context.Load(fields, fs => fs.Include(f => f.InternalName, f => f.FieldTypeKind, f => f.TypeAsString, f => f.ReadOnlyField, f => f.Title));
                        web.Context.ExecuteQueryRetry();

                        var keyColumnType = "Text";
                        var parsedKeyColumn = parser.ParseString(listInstance.DataRows.KeyColumn);
                        if (!string.IsNullOrEmpty(parsedKeyColumn))
                        {
                            var keyColumn = fields.FirstOrDefault(f => f.InternalName.Equals(parsedKeyColumn, StringComparison.InvariantCultureIgnoreCase));
                            if (keyColumn != null)
                            {
                                switch (keyColumn.FieldTypeKind)
                                {
                                    case FieldType.User:
                                    case FieldType.Lookup:
                                        keyColumnType = "Lookup";
                                        break;

                                    case FieldType.URL:
                                        keyColumnType = "Url";
                                        break;

                                    case FieldType.DateTime:
                                        keyColumnType = "DateTime";
                                        break;

                                    case FieldType.Number:
                                    case FieldType.Counter:
                                        keyColumnType = "Number";
                                        break;
                                }
                            }
                        }

                        foreach (var dataRow in listInstance.DataRows)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_list_item__0_, listInstance.DataRows.IndexOf(dataRow) + 1);

                                bool processItem = true;
                                ListItem listitem = null;

                                if (!string.IsNullOrEmpty(listInstance.DataRows.KeyColumn))
                                {
                                    // Get value from key column
                                    var dataRowValues = dataRow.Values.Where(v => v.Key == listInstance.DataRows.KeyColumn).ToList();

                                    // if it is empty, skip the check
                                    if (dataRowValues.Any())
                                    {
                                        var keyColumnValue = parser.ParseString(dataRowValues.FirstOrDefault().Value);
                                        if (keyColumnType == "DateTime")
                                        {
                                            keyColumnValue = DateTime.Parse(keyColumnValue).ToString("s") + "Z";
                                        }
                                        var query = $@"<View><Query><Where><Eq><FieldRef Name=""{parsedKeyColumn}""/><Value {(keyColumnType == "DateTime" ? "IncludeTimeValue='TRUE'" : "")} Type=""{keyColumnType}"">{keyColumnValue}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
                                        var camlQuery = new CamlQuery()
                                        {
                                            ViewXml = query
                                        };
                                        var existingItems = list.GetItems(camlQuery);
                                        list.Context.Load(existingItems);
                                        list.Context.ExecuteQueryRetry();
                                        if (existingItems.Count > 0)
                                        {
                                            if (listInstance.DataRows.UpdateBehavior == UpdateBehavior.Skip)
                                            {
                                                processItem = false;
                                            }
                                            else
                                            {
                                                listitem = existingItems[0];
                                                processItem = true;
                                            }
                                        }
                                    }
                                }

                                if (processItem)
                                {
                                    bool IsNewItem = false;
                                    if (listitem == null)
                                    {
                                        var listitemCI = new ListItemCreationInformation();
                                        listitem = list.AddItem(listitemCI);
                                        IsNewItem = true;
                                    }

                                    ListItemUtilities.UpdateListItem(listitem, parser, dataRow.Values, ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion, IsNewItem);

                                    if (dataRow.Attachments != null && dataRow.Attachments.Count > 0)
                                    {
                                        foreach (var attachment in dataRow.Attachments)
                                        {
                                            attachment.Name = parser.ParseString(attachment.Name);
                                            attachment.Src = parser.ParseString(attachment.Src);
                                            if (!IsNewItem)
                                            {
                                                var overwrite = attachment.Overwrite;
                                                listitem.EnsureProperty(l => l.AttachmentFiles);

                                                Attachment existingItem = null;
                                                if (listitem.AttachmentFiles.Count > 0)
                                                {
                                                    existingItem = listitem.AttachmentFiles.FirstOrDefault(a => a.FileName.Equals(attachment.Name, StringComparison.OrdinalIgnoreCase));
                                                }
                                                if (existingItem != null)
                                                {
                                                    if (overwrite)
                                                    {
                                                        existingItem.DeleteObject();
                                                        web.Context.ExecuteQueryRetry();
                                                        AddAttachment(template, listitem, attachment);
                                                    }
                                                }
                                                else
                                                {
                                                    AddAttachment(template, listitem, attachment);
                                                }
                                            }
                                            else
                                            {
                                                AddAttachment(template, listitem, attachment, IsNewItem);
                                            }
                                        }
                                    }
                                    if (IsNewItem)
                                    {
                                        listitem.Context.Load(listitem, i => i.Id);
                                        listitem.Context.ExecuteQueryRetry();
                                    }
                                    if (dataRow.Security != null && (dataRow.Security.ClearSubscopes || dataRow.Security.CopyRoleAssignments || dataRow.Security.RoleAssignments.Count > 0))
                                    {
                                        listitem.SetSecurity(parser, dataRow.Security, WriteMessage);
                                    }
                                }
                            }
                            catch (ServerException ex)
                            {
                                if (ex.ServerErrorTypeName.Equals("Microsoft.SharePoint.SPDuplicateValuesFoundException", StringComparison.InvariantCultureIgnoreCase)
                                    && applyingInformation.IgnoreDuplicateDataRowErrors)
                                {
                                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_duplicate);
                                    continue;
                                }
                                if (ex.ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase)
                                    && ex.Message.Equals("To add an item to a document library, use SPFileCollection.Add()", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    // somebody tries to add new items to a document library
                                    var warning = string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_notsupported_0, listInstance.Title);
                                    scope.LogWarning(warning);
                                    WriteMessage(warning, ProvisioningMessageType.Warning);
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_failed___0_____1_, ex.Message, ex.StackTrace);
                                throw;
                            }
                        }
                    }
                }


                #endregion DataRows
            }


            return parser;
        }


        private static void AddAttachment(ProvisioningTemplate template, ListItem listitem, Model.SharePoint.InformationArchitecture.DataRowAttachment attachment, bool SkipExecuteQuery = false)
        {
            listitem.AttachmentFiles.AddUsingPath(ResourcePath.FromDecodedUrl(attachment.Name), FileUtilities.GetFileStream(template, attachment.Src));
            if (!SkipExecuteQuery)
                listitem.Context.ExecuteQueryRetry();
            else
                listitem.Update();
        }

        private static bool ShouldNotExtractList(ProvisioningTemplateCreationInformation creationInfo, List siteList)
        {
            if (creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.Lists != null
                && creationInfo.ExtractConfiguration.Lists.HasLists
                &&
                !creationInfo.ExtractConfiguration.Lists.Lists.Any(i =>
                {
                    if (Guid.TryParse(i.Title, out Guid listId))
                    {
                        return (listId == siteList.Id) && i.IncludeItems;
                    }
                    else
                    {
                        return (false);
                    }
                })
                && !creationInfo.ExtractConfiguration.Lists.Lists.Any(i => i.Title.Equals(siteList.Title) && i.IncludeItems)
                && !creationInfo.ExtractConfiguration.Lists.Lists.Any(i => siteList.RootFolder.ServerRelativeUrl.EndsWith(i.Title, StringComparison.InvariantCultureIgnoreCase) && i.IncludeItems))
            {
                return true;
            }

            return false;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var lists = web.Lists;
                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url, w => w.Id);
                web.Context.Load(lists,
                  lc => lc.IncludeWithDefaultProperties(
                        l => l.RootFolder.ServerRelativeUrl,
                        l => l.EnableAttachments,
                        l => l.ContentTypes,
                        l => l.Fields.IncludeWithDefaultProperties(
                          f => f.Id,
                          f => f.Title,
                          f => f.TypeAsString,
                          f => f.ReadOnlyField,
                          f => f.Hidden,
                          f => f.InternalName,
                          f => f.DefaultValue,
                          f => f.Required))
                  );
                web.Context.ExecuteQueryRetry();

                var allLists = new List<List>();

                var listsToProcess = lists.AsEnumerable().Where(l => l.Hidden == false || l.Hidden == creationInfo.IncludeHiddenLists).ToArray();
                foreach (var siteList in listsToProcess)
                {
                    if (ShouldNotExtractList(creationInfo, siteList))
                    {
                        continue;
                    }
                    var extractionConfig = creationInfo.ExtractConfiguration.Lists.Lists.FirstOrDefault(e => e.Title.Equals(siteList.Title) || siteList.RootFolder.ServerRelativeUrl.EndsWith(e.Title, StringComparison.InvariantCultureIgnoreCase));
                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
                    Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig = null;
                    if (extractionConfig.Query != null)
                    {
                        queryConfig = extractionConfig.Query;

                        camlQuery = new CamlQuery();

                        if (string.IsNullOrEmpty(queryConfig.CamlQuery))
                        {
                            queryConfig.CamlQuery = "<Order><FieldRef Name='ID' /></Order>";
                        }
                        string viewXml = $"<View Scope=\"RecursiveAll\"><Query>{queryConfig.CamlQuery}</Query>";
                        if (queryConfig.IncludeAttachments && siteList.EnableAttachments)
                        {
                            if (queryConfig.ViewFields == null)
                            {
                                queryConfig.ViewFields = new List<string>();
                            }
                            else if (!queryConfig.ViewFields.Contains("Attachments"))
                            {
                                queryConfig.ViewFields.Add("Attachments");
                            }
                        }
                        if (queryConfig.ViewFields != null && queryConfig.ViewFields.Count > 0)
                        {
                            viewXml += "<ViewFields>";
                            foreach (var viewField in queryConfig.ViewFields)
                            {
                                viewXml += $"<FieldRef Name='{viewField}' />";
                            }

                            viewXml += "</ViewFields>";
                        }
                        if (queryConfig.RowLimit > 0 || queryConfig.PageSize > 0)
                        {
                            viewXml += $"<RowLimit{(queryConfig.PageSize > 0 ? " Paged=\"TRUE\"" : "")}>{(queryConfig.PageSize > 0 ? queryConfig.PageSize : queryConfig.RowLimit)}</RowLimit>";
                        }
                        viewXml += "</View>";
                        camlQuery.ViewXml = viewXml;

                    }

                    var listInstance = template.Lists.FirstOrDefault(l => siteList.RootFolder.ServerRelativeUrl.Equals(UrlUtility.Combine(web.ServerRelativeUrl, l.Url)));
                    if (listInstance != null)
                    {
                        do
                        {
                            camlQuery.ListItemCollectionPosition = RetrieveItems(web, template, creationInfo, scope, siteList, extractionConfig, camlQuery, queryConfig, listInstance, siteList.ContentTypes[0].Id.StringValue);

                        } while (camlQuery.ListItemCollectionPosition != null);
                    }
                }
            }
            return template;
        }

        private ListItemCollectionPosition RetrieveItems(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, List siteList, Model.Configuration.Lists.Lists.ExtractListsListsConfiguration extractionConfiguration, CamlQuery camlQuery, Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig, ListInstance listInstance, string defaultContentTypeId)
        {
            var items = siteList.GetItems(camlQuery);
            siteList.Context.Load(items, i => i.IncludeWithDefaultProperties(li => li.FieldValuesAsText), i => i.ListItemCollectionPosition);
            if (queryConfig != null && queryConfig.ViewFields != null && queryConfig.ViewFields.Count > 0)
            {
                foreach (var viewField in queryConfig.ViewFields)
                {
                    if (siteList.Fields.FirstOrDefault(f => f.InternalName == viewField) != null)
                    {
                        siteList.Context.Load(items, i => i.Include(li => li[viewField]));
                    }
                }
            }
            siteList.Context.ExecuteQueryRetry();
            var baseUri = new Uri(web.Url);
            if (siteList.BaseType == BaseType.DocumentLibrary)
            {
                ProcessLibraryItems(web, siteList, template, listInstance, extractionConfiguration, queryConfig, creationInfo, scope, items, baseUri, defaultContentTypeId);
            }
            else
            {
                ProcessListItems(web, siteList, listInstance, creationInfo, extractionConfiguration, queryConfig, baseUri, items, scope);
            }
            return items.ListItemCollectionPosition;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Lists.Any(l => l.DataRows.Any());
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.Lists.HasLists;
            }
            return _willExtract.Value;
        }

        private ProvisioningTemplate ProcessLibraryItems(Web web,
            List siteList,
            ProvisioningTemplate template,
            ListInstance listInstance,
            Model.Configuration.Lists.Lists.ExtractListsListsConfiguration extractionConfig,
            Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig,
            ProvisioningTemplateCreationInformation creationInfo,
            PnPMonitoredScope scope,
            ListItemCollection items,
            Uri baseUri,
            string defaultContentTypeId)
        {
            var itemCount = 1;
            var listDefaultColumnValues = siteList.GetDefaultColumnValues();

            foreach (var item in items)
            {
                switch (item.FileSystemObjectType)
                {
                    case FileSystemObjectType.File:
                        {
                            //PnP:File
                            ProcessDocumentRow(web, siteList, baseUri, item, listInstance, template, creationInfo, scope, itemCount, items.Count, defaultContentTypeId);
                            break;
                        }
                    case FileSystemObjectType.Folder:
                        {
                            //PnP:Folder
                            ProcessFolderRow(web, item, siteList, listInstance, queryConfig, listDefaultColumnValues, template, scope);
                            break;
                        }
                    default:
                        {
                            //PnP:DataRow
                            ProcessDataRow(web, siteList, item, listInstance, extractionConfig, queryConfig, baseUri, creationInfo, scope);
                            break;
                        }
                }
                itemCount++;
            }

            //Process Forms Folder 
            ProcessFormsFolder(web, siteList, listInstance, template, scope);
            return template;
        }

        //Export Files referred to in NewDocumentTemplates
        private static void ProcessFormsFolder(Web web, List spList, ListInstance listInstance, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            Microsoft.SharePoint.Client.Folder formsFolder = null;
            try
            {
                web.EnsureProperties(w => w.Url);
                spList.EnsureProperties(l => l.RootFolder.ServerRelativeUrl);
                formsFolder = web.GetFolderByServerRelativeUrl(spList.RootFolder.ServerRelativeUrl + "/Forms");
                web.Context.ExecuteQueryRetry();
            }
            catch (Exception)
            {
                formsFolder = null;
            }
            if (formsFolder != null)
            {
                var baseUri = new Uri(web.Url);

                foreach (var instanceView in listInstance.Views)
                {
                    if (instanceView.SchemaXml.Contains("NewDocumentTemplates"))
                    {
                        var viewSchema = System.Xml.Linq.XDocument.Parse(instanceView.SchemaXml);
                        var templateElement = viewSchema.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "NewDocumentTemplates");
                        if (templateElement != null)
                        {
                            var NewDocumentTemplates = Newtonsoft.Json.Linq.JArray.Parse(templateElement.Value);
                            foreach (var templateFile in NewDocumentTemplates.SelectTokens("..url"))
                            {
                                var FileTemplate = templateFile.Parent.Parent as Newtonsoft.Json.Linq.JObject;
                                if (FileTemplate != null)
                                {
                                    var contentTypeId = FileTemplate["contentTypeId"]?.ToString();
                                    var url = FileTemplate["url"]?.ToString();
                                    if (!string.IsNullOrWhiteSpace(url) && !string.IsNullOrWhiteSpace(contentTypeId))
                                    {
                                        var fullUri = new Uri(baseUri, url.Replace("{site}", baseUri.AbsolutePath.TrimEnd('/')));
                                        var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                                        var fileName = Uri.UnescapeDataString(fullUri.Segments[fullUri.Segments.Length - 1]);

                                        var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart('/');

                                        Microsoft.SharePoint.Client.File myFile = web.GetFileByUrl($"{templateFolderPath}/{fileName}");
                                        web.Context.Load(myFile);
                                        var stream = myFile.OpenBinaryStream();
                                        web.Context.ExecuteQueryRetry();

                                        template.Connector.SaveFileStream(myFile.Name, templateFolderPath, stream.Value);

                                        Model.File newFile = new Model.File()
                                        {
                                            Folder = templateFolderPath,
                                            Src = $"{templateFolderPath}/{fileName}",
                                            TargetFileName = myFile.Name,
                                            Overwrite = true,
                                            Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), myFile.Level.ToString())
                                        };

                                        template.Files.Add(newFile);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ProcessDocumentRow(Web web, List siteList, Uri baseUri, ListItem listItem, ListInstance listInstance, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, int itemCount, int itemsCount, string defaultContentTypeId)
        {
            var myFile = listItem.File; ;
            web.Context.Load(myFile,
                f => f.Name,
                f => f.ServerRelativePath,
                f => f.UniqueId,
                f => f.Level);
            web.Context.ExecuteQueryRetry();

            // If we got here it's a file, let's grab the file's path and name
            var fullUri = new Uri(baseUri, myFile.ServerRelativePath.DecodedUrl);
            var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
            var fileName = Uri.UnescapeDataString(Path.GetFileName(fullUri.AbsoluteUri));

            var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

            WriteSubProgress("Library", $"{listInstance.Title} : {myFile.Name}", itemCount, itemsCount);

            // Avoid duplicate file entries
            Model.File newFile = null;
            bool addFile = false;

            newFile = template.Files.FirstOrDefault(f => f.Src.Equals($"{templateFolderPath}/{fileName}", StringComparison.CurrentCultureIgnoreCase));
            if (newFile == null)
            {

                newFile = new Model.File()
                {
                    Folder = templateFolderPath,
                    Src = $"{templateFolderPath}/{fileName}",
                    Overwrite = true,
                    Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), myFile.Level.ToString())
                };
                addFile = true;
            }

            ExtractFileSettings(web, siteList, myFile.UniqueId, ref newFile, defaultContentTypeId, scope);

            if (addFile && creationInfo.PersistBrandingFiles)
            {
                var file = listItem.File;
                web.Context.Load(file);
                web.Context.ExecuteQueryRetry();
                var spFileStream = file.OpenBinaryStream();
                web.Context.ExecuteQueryRetry();
                using (var streamValue = spFileStream.Value)
                {
                    template.Connector.SaveFileStream(file.Name, templateFolderPath, spFileStream.Value);
                }
                template.Files.Add(newFile);
            }
        }

        private void ExtractFileSettings(Web web, List siteList, Guid fileUniqueId, ref Model.File pnpFile, string defaultContentTypeId, PnPMonitoredScope scope)
        {
            try
            {
                var file = web.GetFileById(fileUniqueId);
                web.Context.Load(file,
                    f => f.Level,
                    f => f.ServerRelativePath,
                    f => f.Properties,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.RoleAssignments,
                    f => f.ListItemAllFields.RoleAssignments.Include(r => r.Member, r => r.RoleDefinitionBindings),
                    f => f.ListItemAllFields.HasUniqueRoleAssignments,
                    f => f.ListItemAllFields.ParentList,
                    f => f.ListItemAllFields.ContentType.StringId);

                web.Context.ExecuteQueryRetry();

                //export PnPFile FieldValues
                if (file.ListItemAllFields.FieldValues.Any())
                {
                    var fieldValues = file.ListItemAllFields.FieldValues;

                    var fieldValuesAsText = file.ListItemAllFields.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

                    #region //**** get correct Content Type
                    string ctId = string.Empty;
                    foreach (var ct in web.ContentTypes.OrderByDescending(c => c.StringId.Length))
                    {
                        if (file.ListItemAllFields.ContentType.StringId.StartsWith(ct.StringId) && file.ListItemAllFields.ContentType.StringId != defaultContentTypeId) // skip if it is the default content type
                        {
                            pnpFile.Properties.Add("ContentTypeId", ct.StringId);
                            break;
                        }
                    }
                    #endregion //**** get correct Content Type

                    foreach (var fieldValue in fieldValues)
                    {
                        if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                        {
                            var field = siteList.Fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);
                            string value = string.Empty;
                            //ignore read only fields
                            if (!field.ReadOnlyField || WriteableReadOnlyField.Contains(field.InternalName.ToLower()))
                            {
                                value = TokenizeValue(web, field.TypeAsString, fieldValue, fieldValuesAsText[field.InternalName]);

                                if (fieldValue.Key == "ContentTypeId" && fieldValue.Key == "Attachments")
                                {
                                    value = null; //it's already in Properties - we can ignore here
                                }
                            }

                            // We process real values only
                            if (value != null && !String.IsNullOrEmpty(value) && value != "[]")
                            {
                                pnpFile.Properties[fieldValue.Key] = value;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of File with uniqueId {0} failed", fileUniqueId);
            }
        }

        private static readonly string[] WriteableReadOnlyField = new[]
        {
            "description","publishingpagelayout", "contenttypeid","bannerimageurl","_originalsourceitemid","_originalsourcelistid","_originalsourcesiteid","_originalsourcewebid","_originalsourceurl"
        };

        private string TokenizeValue(Web web, string fieldTypeAsString, KeyValuePair<string, object> fieldValue, string fieldValueAsText)
        {
            string value = string.Empty;
            switch (fieldTypeAsString)
            {
                case "URL":
                    value = Tokenize(fieldValueAsText, web.Url, web);
                    break;
                case "User":
                    var userFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldUserValue;
                    if (userFieldValue != null)
                    {
                        if (!string.IsNullOrEmpty(userFieldValue.Email))
                            value = userFieldValue.Email;
                        else if (userFieldValue.LookupValue == web.GetEveryoneExceptExternalUsersClaimName())
                            value = "{everyonebutexternalusers}";
                    }
                    break;
                case "UserMulti":
                    var userMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldUserValue[];
                    if (userMultiFieldValue != null)
                    {
                        value = string.Join(",", userMultiFieldValue.Select(u => u.Email).ToArray())?.TrimEnd(new char[] { ',' }).Trim(new char[] { ',' });
                        if (userMultiFieldValue.Any(u => u.LookupValue == web.GetEveryoneExceptExternalUsersClaimName()))
                        {
                            if (!string.IsNullOrEmpty(value))
                                value = value + ",";
                            value = value + "{everyonebutexternalusers}";
                        }
                    }
                    break;
                case "Lookup":
                    var lookupFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldLookupValue;
                    if (lookupFieldValue != null)
                    {
                        value = lookupFieldValue.LookupId.ToString();
                    }
                    break;
                case "LookupMulti":
                    var lookupMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.FieldLookupValue[];
                    if (lookupMultiFieldValue != null)
                    {
                        value = value = string.Join(",", lookupMultiFieldValue.Select(l => l.LookupId).ToArray())?.TrimEnd(new char[] { ',' });
                    }
                    break;
                case "TaxonomyFieldType":
                    var taxonomyFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue;
                    if (taxonomyFieldValue != null)
                    {
                        value = $"{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}";
                    }
                    break;
                case "TaxonomyFieldTypeMulti":
                    var taxonomyMultiFieldValue = fieldValue.Value as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection;
                    if (taxonomyMultiFieldValue != null)
                    {
                        string terms = "";
                        foreach (var term in taxonomyMultiFieldValue)
                        {
                            terms += $"{term.Label}|{term.TermGuid};";
                        }
                        value = terms.TrimEnd(new char[] { ';' });
                    }
                    break;
                case "DateTime":
                    var dateTimeFieldValue = fieldValue.Value as DateTime?;
                    if (dateTimeFieldValue.HasValue)
                    {
                        value = dateTimeFieldValue.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
                    }
                    break;
                case "MultiChoice":
                    if (fieldValue.Value != null)
                    {
                        var multiChoiceArray = fieldValue.Value as string[];
                        value = string.Join(";#", multiChoiceArray.Select(v => Tokenize(v, web.Url, web)).ToList());
                    }
                    break;
                case "ContentTypeIdFieldType":
                default:
                    value = Tokenize(fieldValue.Value?.ToString(), web.Url, web);
                    break;
            }

            return value;
        }

        public Model.Folder ExtractFolderSettings(Web web, List siteList, List<Dictionary<string, string>> listDefaultValues, string serverRelativePathToFolder, PnPMonitoredScope scope, Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig)
        {
            Model.Folder pnpFolder = null;
            try
            {
                Microsoft.SharePoint.Client.Folder spFolder = web.GetFolderByServerRelativeUrl(serverRelativePathToFolder);
                web.Context.Load(spFolder,
                    f => f.Name,
                    f => f.ServerRelativeUrl,
                    f => f.Properties,
                    f => f.ListItemAllFields,
                    f => f.ListItemAllFields.RoleAssignments,
                    f => f.ListItemAllFields.RoleAssignments.Include(r => r.Member, r => r.RoleDefinitionBindings),
                    f => f.ListItemAllFields.HasUniqueRoleAssignments,
                    f => f.ListItemAllFields.ParentList,
                    f => f.ListItemAllFields.ContentType.StringId);
                web.Context.Load(web,
                    w => w.AssociatedOwnerGroup,
                    w => w.AssociatedMemberGroup,
                    w => w.AssociatedVisitorGroup,
                    w => w.Title,
                    w => w.Url,
                    w => w.RoleDefinitions.Include(r => r.RoleTypeKind, r => r.Name),
                    w => w.ContentTypes.Include(c => c.Id, c => c.Name, c => c.StringId));
                web.Context.ExecuteQueryRetry();

                pnpFolder = new Model.Folder(spFolder.Name);

                //export PnPFolder Properties
                if (spFolder.Properties.FieldValues.Any())
                {
                    foreach (var propKey in spFolder.Properties.FieldValues.Keys.Where(k => !k.StartsWith("vti_") && !k.StartsWith("docset_")))
                    {
                        pnpFolder.PropertyBagEntries.Add(new PropertyBagEntry() { Key = propKey, Value = spFolder.Properties.FieldValues[propKey].ToString() });
                    }
                }

                //export PnPFolder FieldValues
                if (spFolder.ListItemAllFields.FieldValues.Any())
                {
                    var list = spFolder.ListItemAllFields.ParentList;

                    var fields = list.Fields;
                    web.Context.Load(fields, fs => fs.IncludeWithDefaultProperties(f => f.TypeAsString, f => f.InternalName, f => f.Title));
                    web.Context.ExecuteQueryRetry();

                    var fieldValues = spFolder.ListItemAllFields.FieldValues;

                    var fieldValuesAsText = spFolder.ListItemAllFields.EnsureProperty(li => li.FieldValuesAsText).FieldValues;

                    #region //**** get correct Content Type
                    string ctId = string.Empty;
                    foreach (var ct in web.ContentTypes.OrderByDescending(c => c.StringId.Length))
                    {
                        if (spFolder.ListItemAllFields.ContentType.StringId.StartsWith(ct.StringId))
                        {
                            pnpFolder.ContentTypeID = ct.StringId;
                            break;
                        }
                    }
                    #endregion //**** get correct Content Type

                    var filteredFieldValues = fieldValues.ToList();
                    if (queryConfig != null && queryConfig.ViewFields != null && queryConfig.ViewFields.Count > 0)
                    {
                        filteredFieldValues = fieldValues.Where(f => queryConfig.ViewFields.Contains(f.Key)).ToList();
                    }
                    foreach (var fieldValue in filteredFieldValues)
                    {
                        if (fieldValue.Value != null && !string.IsNullOrEmpty(fieldValue.Value.ToString()))
                        {
                            var field = siteList.Fields.FirstOrDefault(fs => fs.InternalName == fieldValue.Key);
                            string value = string.Empty;

                            //ignore read only fields
                            if (!field.ReadOnlyField || WriteableReadOnlyField.Contains(field.InternalName.ToLower()))
                            {
                                value = TokenizeValue(web, field.TypeAsString, fieldValue, fieldValuesAsText[field.InternalName]);
                            }
                            
                            //We process moderation status, ideally this shoud be managed with a new attribute in Folder, but it requires a new schema version
                            if (fieldValue.Key.Equals("_ModerationStatus", StringComparison.InvariantCultureIgnoreCase))
                            {
                                value = TokenizeValue(web, field.TypeAsString, fieldValue, fieldValuesAsText[field.InternalName]);
                            }

                            if (fieldValue.Key.Equals("ContentTypeId", StringComparison.InvariantCultureIgnoreCase) || fieldValue.Key.Equals("Attachments", StringComparison.InvariantCultureIgnoreCase))
                            {
                                value = null; //ignore here since already in dataRow
                            }

                            if (fieldValue.Key.Equals("HTML_x0020_File_x0020_Type", StringComparison.CurrentCultureIgnoreCase) &&
                                fieldValuesAsText["HTML_x0020_File_x0020_Type"] == "OneNote.Notebook")
                            {
                                pnpFolder.Properties.Add("File_x0020_Type", "OneNote.Notebook");
                                pnpFolder.Properties.Add(fieldValue.Key, "OneNote.Notebook");
                                value = null;
                            }

                            // We process real values only
                            if (!string.IsNullOrWhiteSpace(value) && value != "[]")
                            {
                                pnpFolder.Properties.Add(fieldValue.Key, value);
                            }
                        }
                    }
                }

                //export PnPFolder default values
                if (listDefaultValues != null)
                {
                    var href = Uri.UnescapeDataString(spFolder.ServerRelativeUrl);
                    href = href.Replace(siteList.RootFolder.ServerRelativeUrl, "/").Replace("//", "/");

                    var defaultValues = listDefaultValues.Where(dv => dv["Path"] == href);
                    foreach (var defaultValue in defaultValues)
                    {
                        pnpFolder.DefaultColumnValues.Add(defaultValue["Field"], defaultValue["Value"]);
                    }
                }
            }
            catch (Exception ex)
            {
                scope.LogError(ex, "Extract of Folder {0} failed", serverRelativePathToFolder);
            }
            return pnpFolder;
        }

        private void ProcessFolderRow(Web web, ListItem listItem, List siteList, ListInstance listInstance, Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig, List<Dictionary<string, string>> listDefaultValues, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            listItem.EnsureProperties(it => it.ParentList.RootFolder.ServerRelativeUrl);
            string serverRelativeListUrl = listItem.ParentList.RootFolder.ServerRelativeUrl;
            string folderPath = listItem.FieldValuesAsText["FileRef"].Substring(serverRelativeListUrl.Length).TrimStart(new char[] { '/' });

            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                //listItem.EnsureProperties(it => it.Folder.UniqueId);
                string[] folderSegments = folderPath.Split('/');
                Model.Folder pnpFolder = null;
                for (int i = 0; i < folderSegments.Length; i++)
                {
                    if (i == 0)
                    {
                        pnpFolder = listInstance.Folders.FirstOrDefault(f => f.Name.Equals(folderSegments[i], StringComparison.CurrentCultureIgnoreCase));
                        if (pnpFolder == null)
                        {
                            string pathToCurrentFolder = string.Format("{0}/{1}", serverRelativeListUrl, string.Join("/", folderSegments.Take(i + 1)));
                            pnpFolder = ExtractFolderSettings(web, siteList, listDefaultValues, pathToCurrentFolder, scope, queryConfig);
                            listInstance.Folders.Add(pnpFolder);
                        }
                    }
                    else
                    {
                        var childFolder = pnpFolder.Folders.FirstOrDefault(f => f.Name.Equals(folderSegments[i], StringComparison.CurrentCultureIgnoreCase));
                        if (childFolder == null)
                        {
                            string pathToCurrentFolder = string.Format("{0}/{1}", serverRelativeListUrl, string.Join("/", folderSegments.Take(i + 1)));
                            childFolder = ExtractFolderSettings(web, siteList, listDefaultValues, pathToCurrentFolder, scope, queryConfig);
                            pnpFolder.Folders.Add(childFolder);
                        }
                        pnpFolder = childFolder;
                    }
                }
            }
        }

        private ListInstance ProcessListItems(Web web,
            List siteList,
            ListInstance listInstance,
            ProvisioningTemplateCreationInformation creationInfo,
            Model.Configuration.Lists.Lists.ExtractListsListsConfiguration extractionConfig,
            Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig,
            Uri baseUri,
            ListItemCollection items,
            PnPMonitoredScope scope)
        {
            if (!string.IsNullOrEmpty(extractionConfig.KeyColumn))
            {
                listInstance.DataRows.KeyColumn = extractionConfig.KeyColumn;
                listInstance.DataRows.UpdateBehavior = extractionConfig.UpdateBehavior;
            }

            var itemCount = 1;
            foreach (var item in items)
            {
                WriteSubProgress("List", listInstance.Title, itemCount, items.Count);

                var dataRow = ProcessDataRow(web, siteList, item, listInstance, extractionConfig, queryConfig, baseUri, creationInfo, scope);

                listInstance.DataRows.Add(dataRow);
                itemCount++;
            }
            return listInstance;
        }

        private Model.DataRow ProcessDataRow(Web web, List siteList, ListItem item, ListInstance listInstance, Model.Configuration.Lists.Lists.ExtractListsListsConfiguration extractionConfig, Model.Configuration.Lists.Lists.ExtractListsQueryConfiguration queryConfig, Uri baseUri, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var dataRow = new Model.DataRow();
            var filteredFieldValues = item.FieldValues.ToList();
            if (queryConfig != null && queryConfig.ViewFields != null && queryConfig.ViewFields.Count > 0)
            {
                filteredFieldValues = item.FieldValues.Where(f => queryConfig.ViewFields.Contains(f.Key)).ToList();
            }
            if (queryConfig != null && queryConfig.IncludeAttachments)
            {
                filteredFieldValues = filteredFieldValues.Where(f => !f.Key.Equals("Attachments", StringComparison.InvariantCultureIgnoreCase)).ToList();
            }
            foreach (var fieldValue in filteredFieldValues)
            {
                var value = item.FieldValuesAsText[fieldValue.Key];//FieldValuesAsText strips of html and returns empty string in case all info is in attributes like for canvascontrol in HostedAppsConfig
                var skip = extractionConfig.SkipEmptyFields && item[fieldValue.Key]==null;
                if (!skip)
                {
                    string parsedValue = TokenizeValue(web, siteList.Fields.FirstOrDefault(f => f.InternalName == fieldValue.Key).TypeAsString, fieldValue, value);
                    if(!(extractionConfig.SkipEmptyFields && string.IsNullOrEmpty(parsedValue)))
                        dataRow.Values.Add(fieldValue.Key, parsedValue);
                }
            }
            if (queryConfig != null && queryConfig.IncludeAttachments && siteList.EnableAttachments && (bool)item["Attachments"])
            {
                item.Context.ExecuteQueryRetry();
                item.EnsureProperty(i => i.AttachmentFiles);
                foreach (var attachmentFile in item.AttachmentFiles)
                {
                    var fullUri = new Uri(baseUri, attachmentFile.ServerRelativePath.DecodedUrl);
                    var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                    var targetFolder = $"ListData/SITE_{web.Id.ToString("N")}/LIST_{siteList.Id.ToString("N")}/Attachments/{item.Id}";
                    dataRow.Attachments.Add(new Model.SharePoint.InformationArchitecture.DataRowAttachment()
                    {
                        Name = attachmentFile.FileNameAsPath.DecodedUrl,
                        Src = $"{targetFolder}/{attachmentFile.FileNameAsPath.DecodedUrl}"
                    });
                    if (creationInfo.PersistBrandingFiles)
                    {
                        PersistFile(web, creationInfo, scope, attachmentFile.ServerRelativePath.DecodedUrl, attachmentFile.FileNameAsPath.DecodedUrl, targetFolder);
                    }
                }
            }
            return dataRow;
        }

        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string fileServerRelativeUrl, string targetFilename, string targetFolder)
        {
            if (creationInfo.FileConnector != null)
            {
                var targetContainer = Uri.UnescapeDataString(targetFolder).Trim('/').Replace("/", "\\");

                using (Stream s = GetAttachmentStream((ClientContext)web.Context, fileServerRelativeUrl))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(
                            targetFilename, targetContainer, s);
                    }
                }
            }
            else
            {
                scope.LogError($"No connector present to persist file {fileServerRelativeUrl}");
            }
        }

        private static Stream GetAttachmentStream(ClientContext context, string fileServerRelativeUrl)
        {
            try
            {
                var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
                context.Load(file);
                context.ExecuteQueryRetry();
                if (file.Exists)
                {
                    MemoryStream stream = new MemoryStream();
                    var streamResult = file.OpenBinaryStream();
                    context.ExecuteQueryRetry();

                    streamResult.Value.CopyTo(stream);

                    // Set the stream position to the beginning
                    stream.Position = 0;
                    return stream;
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, $"Attachment file {fileServerRelativeUrl} not found");
#pragma warning disable CA2200
                throw ex;
#pragma warning restore CA2200
            }
            return null;
        }
    }
}
