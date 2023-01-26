﻿using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Diagnostics;
using PnP.Framework.Enums;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Extensions;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using Field = PnP.Framework.Provisioning.Model.Field;
using SPField = Microsoft.SharePoint.Client.Field;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectField : ObjectHandlerBase
    {
        private readonly FieldAndListProvisioningStepHelper.Step _step;

        public override string Name
        {
#if DEBUG
            get { return $"Fields ({_step})"; }
#else
            get { return $"Fields"; }
#endif
        }

        public override string InternalName => "Fields";

        public ObjectField(FieldAndListProvisioningStepHelper.Step step)
        {
            this._step = step;
        }

        public class DuplicateKeyComparer<TKey>
            :
                IComparer<TKey> where TKey : IComparable
        {
            #region IComparer<TKey> Members

            public int Compare(TKey x, TKey y)
            {
                int result = x.CompareTo(y);
                return result == 0 ? -1 : result;
            }

            #endregion
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not provisioning fields. Technically this can be done but it's not a recommended practice
                if (web.IsSubSite() && !applyingInformation.ProvisionFieldsToSubWebs)
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Context_web_is_subweb__skipping_site_columns);
                    WriteMessage("This template contains fields and you are provisioning to a subweb. If you still want to provision these fields, set the ProvisionFieldsToSubWebs property to true.", ProvisioningMessageType.Warning);
                    return parser;
                }

                var existingFields = web.Fields;

                web.Context.Load(existingFields, fs => fs.Include(f => f.Id));
                web.Context.ExecuteQueryRetry();
                var existingFieldIds = existingFields.AsEnumerable<SPField>().Select(l => l.Id).ToList();

                SortedList<string, Field> fieldDict = new SortedList<string, Field>(new DuplicateKeyComparer<string>());
                foreach (Field siteField in template.SiteFields)
                {
                    var step = siteField.GetFieldProvisioningStep(parser);

                    if (step == _step)
                    {
                        var fieldRef = (string)XElement.Parse(parser.ParseXmlString(siteField.SchemaXml)).Attribute("FieldRef") + "";
                        fieldDict.Add(fieldRef, siteField);
                    }
                }

                var fields = fieldDict.Values.ToList();

                var currentFieldIndex = 0;
                foreach (var field in fields)
                {
                    currentFieldIndex++;
                    var fieldSchemaElement = XElement.Parse(parser.ParseXmlString(field.SchemaXml));
                    var fieldId = fieldSchemaElement.Attribute("ID").Value;
                    var fieldInternalName = fieldSchemaElement.Attribute("InternalName")?.Value ?? fieldSchemaElement.Attribute("Name")?.Value;
                    WriteSubProgress("Field", !string.IsNullOrWhiteSpace(fieldInternalName) ? fieldInternalName : fieldId, currentFieldIndex, fields.Count);
                    if (!existingFieldIds.Contains(Guid.Parse(fieldId)))
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Adding_field__0__to_site, fieldId);
                            CreateField(web, fieldSchemaElement, scope, parser, field.SchemaXml);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Fields_Adding_field__0__failed___1_____2_, fieldId, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                    else
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__in_site, fieldId);
                            UpdateField(web, fieldId, fieldSchemaElement, scope, parser, field.SchemaXml);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__failed___1_____2_, fieldId, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                }
            }
            WriteMessage($"Done processing fields", ProvisioningMessageType.Completed);
            return parser;
        }

        private void UpdateField(Web web, string fieldId, XElement templateFieldElement, PnPMonitoredScope scope, TokenParser parser, string originalFieldXml)
        {
            var existingField = web.Fields.GetById(Guid.Parse(fieldId));
            web.Context.Load(existingField, f => f.SchemaXml);
            web.Context.ExecuteQueryRetry();

            XElement existingFieldElement = XElement.Parse(existingField.SchemaXml);

            XNodeEqualityComparer equalityComparer = new XNodeEqualityComparer();

            if (equalityComparer.GetHashCode(existingFieldElement) != equalityComparer.GetHashCode(templateFieldElement)) // Is field different in template?
            {
                if (existingFieldElement.Attribute("Type").Value == templateFieldElement.Attribute("Type").Value) // Is existing field of the same type?
                {
                    var listIdentifier = templateFieldElement.Attribute("List") != null ? templateFieldElement.Attribute("List").Value : null;

                    if (listIdentifier != null)
                    {
                        // Temporary remove list attribute from list
                        templateFieldElement.Attribute("List").Remove();
                    }

                    if (IsFieldXmlValid(parser.ParseXmlString(originalFieldXml), parser, web.Context))
                    {
                        foreach (var attribute in templateFieldElement.Attributes())
                        {
                            if (existingFieldElement.Attribute(attribute.Name) != null)
                            {
                                existingFieldElement.Attribute(attribute.Name).Value = attribute.Value;
                            }
                            else
                            {
                                existingFieldElement.Add(attribute);
                            }
                        }
                        foreach (var element in templateFieldElement.Elements())
                        {
                            if (existingFieldElement.Element(element.Name) != null)
                            {
                                existingFieldElement.Element(element.Name).Remove();
                            }
                            existingFieldElement.Add(element);
                        }

                        if (string.Equals(templateFieldElement.Attribute("Type").Value, "Calculated", StringComparison.OrdinalIgnoreCase))
                        {
                            var fieldRefsElement = existingFieldElement.Descendants("FieldRefs").FirstOrDefault();
                            if (fieldRefsElement != null)
                            {
                                fieldRefsElement.Remove();
                            }
                        }

                        if (existingFieldElement.Attribute("Version") != null)
                        {
                            existingFieldElement.Attributes("Version").Remove();
                        }
                        existingField.SchemaXml = parser.ParseXmlString(existingFieldElement.ToString());
                        existingField.UpdateAndPushChanges(true);
                        web.Context.Load(existingField, f => f.TypeAsString, f => f.DefaultValue);
                        try
                        {
                            web.Context.ExecuteQueryRetry();
                        }
                        catch (Exception ex)
                        {
                            if ((ex is ServerException && (ex as ServerException).ServerErrorTypeName == "Microsoft.SharePoint.Client.ClientServiceTimeoutException")
                               || (ex is WebException && (ex as WebException).Status == WebExceptionStatus.Timeout))
                            {
                                string fieldName = existingFieldElement.Attribute("Name") != null ? existingFieldElement.Attribute("Name").Value : existingFieldElement.Attribute("StaticName").Value;
                                WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__timeout, fieldName), ProvisioningMessageType.Warning);
                                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__timeout, fieldName);

                                web.Context.Load(existingField, f => f.TypeAsString, f => f.DefaultValue);
                                web.Context.ExecuteQueryRetry();
                            }
                            else
                                throw;
                        }

                        bool isDirty = false;
                        if (originalFieldXml.ContainsResourceToken())
                        {
                            var originalFieldElement = XElement.Parse(originalFieldXml);
                            var nameAttributeValue = originalFieldElement.Attribute("DisplayName") != null ? originalFieldElement.Attribute("DisplayName").Value : "";
                            if (nameAttributeValue.ContainsResourceToken())
                            {
                                existingField.TitleResource.SetUserResourceValue(nameAttributeValue, parser);
                                isDirty = true;
                            }
                            var descriptionAttributeValue = originalFieldElement.Attribute("Description") != null ? originalFieldElement.Attribute("Description").Value : "";
                            if (descriptionAttributeValue.ContainsResourceToken())
                            {
                                existingField.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser);
                                isDirty = true;
                            }
                        }

                        if (isDirty)
                        {
                            existingField.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                        if ((existingField.TypeAsString == "TaxonomyFieldType" || existingField.TypeAsString == "TaxonomyFieldTypeMulti"))
                        {
                            var taxField = web.Context.CastTo<TaxonomyField>(existingField);
                            web.Context.Load(taxField);
                            web.Context.ExecuteQueryRetry();

                            if (!string.IsNullOrEmpty(existingField.DefaultValue))
                            {
                                ValidateTaxonomyFieldDefaultValue(taxField);
                            }
                            UpdateTaxonomyField(taxField, existingFieldElement);
                        }
                    }
                    else
                    {
                        // The field Xml was found invalid
                        var tokenString = parser.GetLeftOverTokens(originalFieldXml).Aggregate(String.Empty, (acc, i) => acc + " " + i);
                        scope.LogError("The field was found invalid: {0}", tokenString);
                        throw new Exception($"The field was found invalid: {tokenString}");
                    }
                }
                else
                {
                    var fieldName = existingFieldElement.Attribute("Name") != null ? existingFieldElement.Attribute("Name").Value : existingFieldElement.Attribute("StaticName").Value;
                    WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_Fields_Field__0____1___exists_but_is_of_different_type__Skipping_field_, fieldName, fieldId), ProvisioningMessageType.Warning);
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Fields_Field__0____1___exists_but_is_of_different_type__Skipping_field_, fieldName, fieldId);
                }
            }
        }

        /// <summary>
        /// Tokenizes calculated fieldXml to use tokens for field references
        /// </summary>
        /// <param name="fields">the field collection that the field is contained within</param>
        /// <param name="field">the field to tokenize</param>
        /// <param name="fieldXml">the xml to tokenize</param>
        /// <returns></returns>
        internal static string TokenizeFieldFormula(Microsoft.SharePoint.Client.FieldCollection fields, FieldCalculated field, string fieldXml)
        {
            var schemaElement = XElement.Parse(fieldXml);

            var formulaElement = schemaElement.Descendants("Formula").FirstOrDefault();

            if (formulaElement != null)
            {
                var formulastring = formulaElement.Value;

                if (formulastring != null)
                {
                    var fieldInternalNames = schemaElement.Descendants("FieldRef").Select(fr => fr.Attribute("Name").Value).Distinct();
                    foreach (var fieldInternalName in fieldInternalNames)
                    {
                        var referencedField = fields.GetFieldByInternalName(fieldInternalName);
                        formulastring = formulastring.Replace($"{fieldInternalName}", $"[{referencedField.Title}]");
                    }
                    var fieldRefParent = schemaElement.Descendants("FieldRefs");
                    fieldRefParent.Remove();

                    formulaElement.Value = formulastring;
                }
            }
            return schemaElement.ToString();
        }

        /// <summary>
        /// Replace Field Internal name by Display Name in the Validation formula
        /// (due to a SP issue that when provisioning the field, is expecting the Display name)
        /// https://github.com/SharePoint/PnP-Sites-Core/issues/849
        /// </summary>
        /// <param name="field"></param>
        /// <param name="schemaXml"></param>
        /// <returns></returns>
        internal static string TokenizeFieldValidationFormula(SPField field, string schemaXml)
        {
            var schemaElement = XElement.Parse(field.SchemaXml);

            var validationNode = schemaElement.Elements("Validation").FirstOrDefault();
            if (validationNode != null)
            {
                var validationNodeValue = validationNode.Value;
                validationNode.Value = validationNodeValue.Replace(field.InternalName, string.Format("[{0}]", field.Title));
            }

            return schemaElement.ToString();
        }

        private static string ParseFieldSchema(string schemaXml, Web web, ListCollection lists)
        {
            foreach (var list in lists)
            {
                schemaXml = Regex.Replace(schemaXml, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
            }
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("B"), "{{siteid}}", RegexOptions.IgnoreCase);
            schemaXml = Regex.Replace(schemaXml, web.Id.ToString("D"), "{siteid}", RegexOptions.IgnoreCase);

            return schemaXml;
        }

        private static void CreateField(Web web, XElement templateFieldElement, PnPMonitoredScope scope, TokenParser parser, string originalFieldXml)
        {
            var fieldXml = parser.ParseXmlString(templateFieldElement.ToString());

            if (IsFieldXmlValid(fieldXml, parser, web.Context))
            {
                fieldXml = FieldUtilities.FixLookupField(fieldXml, web);

                var field = web.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddFieldInternalNameHint);
                web.Context.Load(field, f => f.Id, f => f.TypeAsString, f => f.DefaultValue, f => f.InternalName, f => f.Title);
                web.Context.ExecuteQueryRetry();

                // Add newly created field to token set, this allows to create a field + use it in a formula in the same provisioning template
                parser.AddToken(new FieldTitleToken(web, field.InternalName, field.Title));
                parser.AddToken(new FieldIdToken(web, field.InternalName, field.Id));

                bool isDirty = false;

                if (originalFieldXml.ContainsResourceToken())
                {
                    var originalFieldElement = XElement.Parse(originalFieldXml);
                    var nameAttributeValue = originalFieldElement.Attribute("DisplayName") != null ? originalFieldElement.Attribute("DisplayName").Value : "";
                    if (nameAttributeValue.ContainsResourceToken())
                    {
                        field.TitleResource.SetUserResourceValue(nameAttributeValue, parser);
                        isDirty = true;
                    }
                    var descriptionAttributeValue = originalFieldElement.Attribute("Description") != null ? originalFieldElement.Attribute("Description").Value : "";
                    if (descriptionAttributeValue.ContainsResourceToken())
                    {
                        field.DescriptionResource.SetUserResourceValue(descriptionAttributeValue, parser);
                        isDirty = true;
                    }
                }

                if (isDirty)
                {
                    field.Update();
                    web.Context.ExecuteQueryRetry();
                }

                if (field.TypeAsString == "TaxonomyFieldType" || field.TypeAsString == "TaxonomyFieldTypeMulti")
                {
                    var taxField = web.Context.CastTo<TaxonomyField>(field);
                    if (!string.IsNullOrEmpty(field.DefaultValue))
                    {
                        ValidateTaxonomyFieldDefaultValue(taxField);
                    }
                }
            }
            else
            {
                // The field Xml was found invalid
                var tokenString = parser.GetLeftOverTokens(fieldXml).Aggregate(String.Empty, (acc, i) => acc + " " + i);
                scope.LogError("The field was found invalid: {0}", tokenString);
                throw new Exception($"The field was found invalid: {tokenString}");
            }
        }

        private static void UpdateTaxonomyField(TaxonomyField field, XElement taxonomyFieldElement)
        {
            bool isDirty = false;

            var sspIdElement = taxonomyFieldElement.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'SspId']/Value");
            if (sspIdElement != null && Guid.TryParse(sspIdElement.Value, out Guid sspIdValue) && field.SspId.Equals(sspIdValue) == false)
            {
                field.SspId = sspIdValue;
                isDirty = true;
            }

            var termSetIdElement = taxonomyFieldElement.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'TermSetId']/Value");
            if (termSetIdElement != null && Guid.TryParse(termSetIdElement.Value, out Guid termSetIdValue) && field.TermSetId.Equals(termSetIdValue) == false)
            {
                field.TermSetId = termSetIdValue;
                isDirty = true;
            }

            var anchorIdElement = taxonomyFieldElement.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'AnchorId']/Value");
            if (anchorIdElement != null && Guid.TryParse(anchorIdElement.Value, out Guid anchorIdValue) && field.AnchorId.Equals(anchorIdValue) == false)
            {
                field.AnchorId = anchorIdValue;
                isDirty = true;
            }

            var openElement = taxonomyFieldElement.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'Open']/Value");
            if (openElement != null && bool.TryParse(openElement.Value, out bool openValue) && field.Open.Equals(openValue) == false)
            {
                field.Open = openValue;
                isDirty = true;
            }

            var isPathRenderedElement = taxonomyFieldElement.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'IsPathRendered']/Value");
            if (isPathRenderedElement != null && bool.TryParse(isPathRenderedElement.Value, out bool isPathRenderedValue) && field.IsPathRendered.Equals(isPathRenderedValue) == false)
            {
                field.IsPathRendered = isPathRenderedValue;
                isDirty = true;
            }

            if (isDirty)
            {
                field.UpdateAndPushChanges(true);
                field.Context.ExecuteQueryRetry();
            }
        }

        private static void ValidateTaxonomyFieldDefaultValue(TaxonomyField field)
        {
            //get validated value with correct WssIds
            var validatedValue = GetTaxonomyFieldValidatedValue(field, field.DefaultValue);
            if (!string.IsNullOrEmpty(validatedValue) && field.DefaultValue != validatedValue)
            {
                field.DefaultValue = validatedValue;
                field.UpdateAndPushChanges(true);
                field.Context.ExecuteQueryRetry();
            }
        }

        internal static string GetTaxonomyFieldValidatedValue(TaxonomyField field, string defaultValue)
        {
            string res = null;
            object parsedValue = null;
            field.EnsureProperty(f => f.AllowMultipleValues);
            if (field.AllowMultipleValues)
            {
                parsedValue = new TaxonomyFieldValueCollection(field.Context, defaultValue, field);
            }
            else
            {
                TaxonomyFieldValue taxValue = null;
                if (TryParseTaxonomyFieldValue(defaultValue, out taxValue))
                {
                    parsedValue = taxValue;
                }
            }
            if (parsedValue != null)
            {
                var validateValue = field.GetValidatedString(parsedValue);
                field.Context.ExecuteQueryRetry();
                res = validateValue.Value;
            }
            return res;
        }

        private static bool TryParseTaxonomyFieldValue(string value, out TaxonomyFieldValue taxValue)
        {
            bool res = false;
            taxValue = new TaxonomyFieldValue();
            if (!string.IsNullOrEmpty(value))
            {
                string[] split = value.Split(new string[] { ";#" }, StringSplitOptions.None);
                int wssId = 0;

                if (split.Length > 0 && int.TryParse(split[0], out wssId))
                {
                    taxValue.WssId = wssId;
                    res = true;
                }

                if (res && split.Length == 2)
                {
                    var term = split[1];
                    string[] splitTerm = term.Split(new string[] { "|" }, StringSplitOptions.None);
                    Guid termId = Guid.Empty;
                    if (splitTerm.Length > 0)
                    {
                        res = Guid.TryParse(splitTerm[splitTerm.Length - 1], out termId);
                        taxValue.TermGuid = termId.ToString();
                        if (res && splitTerm.Length > 1)
                        {
                            taxValue.Label = splitTerm[0];
                        }
                    }
                    else
                    {
                        res = false;
                    }
                    res = true;
                }
                else if (split.Length == 1 && int.TryParse(value, out wssId))
                {
                    taxValue.WssId = wssId;
                    res = true;
                }
            }
            return res;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not creating field entities.
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Context_web_is_subweb__skipping_site_columns);
                    return template;
                }

                var existingFields = web.Fields;
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.Load(existingFields, fs => fs.Include(f => f.Id, f => f.SchemaXml, f => f.TypeAsString, f => f.InternalName, f => f.Title));
                web.Context.Load(web.Lists, ls => ls.Include(l => l.Id, l => l.Title));
                web.Context.ExecuteQueryRetry();

                var taxTextFieldsToMoveUp = new List<Guid>();
                var calculatedFieldsToMoveDown = new List<Guid>();

                var currentFieldIndex = 0;
                var fieldsToProcessCount = existingFields.Count;
                foreach (var field in existingFields)
                {
                    currentFieldIndex++;
                    WriteSubProgress("Field", field.InternalName, currentFieldIndex, fieldsToProcessCount);
                    if (!BuiltInFieldId.Contains(field.Id))
                    {
                        var fieldXml = field.SchemaXml;
                        XElement element = XElement.Parse(fieldXml);

                        // Check if the field contains a reference to a list. If by Guid, rewrite the value of the attribute to use web relative paths
                        var listIdentifier = element.Attribute("List") != null ? element.Attribute("List").Value : null;
                        if (!string.IsNullOrEmpty(listIdentifier))
                        {
                            //var listGuid = Guid.Empty;
                            fieldXml = ParseFieldSchema(fieldXml, web, web.Lists);
                            element = XElement.Parse(fieldXml);
                        }

                        // Check if the field is of type TaxonomyField
                        if (field.TypeAsString.StartsWith("TaxonomyField"))
                        {
                            var taxField = (TaxonomyField)field;
                            web.Context.Load(taxField, tf => tf.TextField, tf => tf.Id);
                            web.Context.ExecuteQueryRetry();
                            taxTextFieldsToMoveUp.Add(taxField.TextField);

                            fieldXml = TokenizeTaxonomyField(web, element);
                        }

                        // Check if we have version attribute. Remove if exists
                        if (element.Attribute("Version") != null)
                        {
                            element.Attributes("Version").Remove();
                            fieldXml = element.ToString();
                        }
                        if (element.Attribute("Type").Value == "Calculated")
                        {
                            fieldXml = TokenizeFieldFormula(web.Fields, (FieldCalculated)field, fieldXml);
                            calculatedFieldsToMoveDown.Add(field.Id);
                        }
                        if (creationInfo.PersistMultiLanguageResources)
                        {

                            // only persist language values for fields we actually will keep...no point in spending time on this is we clean the field afterwards
                            bool persistLanguages = true;
                            if (creationInfo.BaseTemplate != null)
                            {
                                int index = creationInfo.BaseTemplate.SiteFields.FindIndex(f => Guid.Parse(XElement.Parse(f.SchemaXml).Attribute("ID").Value).Equals(field.Id));

                                if (index > -1)
                                {
                                    persistLanguages = false;
                                }
                            }

                            if (persistLanguages)
                            {
                                var fieldElement = XElement.Parse(fieldXml);
                                var escapedFieldTitle = field.Title.Replace(" ", "_");
                                if (UserResourceExtensions.PersistResourceValue(field.TitleResource, $"Field_{escapedFieldTitle}_DisplayName", template, creationInfo))
                                {
                                    var fieldTitle = $"{{res:Field_{escapedFieldTitle}_DisplayName}}";
                                    fieldElement.SetAttributeValue("DisplayName", fieldTitle);
                                }
                                if (UserResourceExtensions.PersistResourceValue(field.DescriptionResource, $"Field_{escapedFieldTitle}_Description", template, creationInfo))
                                {
                                    var fieldDescription = $"{{res:Field_{escapedFieldTitle}_Description}}";
                                    fieldElement.SetAttributeValue("Description", fieldDescription);
                                }

                                fieldXml = fieldElement.ToString();
                            }

                        }

                        template.SiteFields.Add(new Field() { SchemaXml = fieldXml });
                    }
                }
                // move hidden taxonomy text fields to the top of the list
                foreach (var textFieldId in taxTextFieldsToMoveUp)
                {
                    var field = template.SiteFields.First(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(textFieldId));
                    template.SiteFields.RemoveAll(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(textFieldId));
                    template.SiteFields.Insert(0, field);
                }
                // move calculated fields to the bottom of the list
                // this will not be sufficient in the case of a calculated field is referencing another calculated field
                foreach (var calculatedFieldId in calculatedFieldsToMoveDown)
                {
                    var field = template.SiteFields.First(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(calculatedFieldId));
                    template.SiteFields.RemoveAll(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(calculatedFieldId));
                    template.SiteFields.Add(field);
                }
                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            WriteMessage($"Done processing fields", ProvisioningMessageType.Completed);

            return template;
        }

        private static ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            foreach (var field in baseTemplate.SiteFields)
            {
                XDocument xDoc = XDocument.Parse(field.SchemaXml);
                var id = xDoc.Root.Attribute("ID") != null ? xDoc.Root.Attribute("ID").Value : null;
                if (id != null)
                {
                    int index = template.SiteFields.FindIndex(f => Guid.Parse(XElement.Parse(f.SchemaXml).Attribute("ID").Value).Equals(Guid.Parse(id)));

                    if (index > -1)
                    {
                        template.SiteFields.RemoveAt(index);
                    }
                }
            }

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.SiteFields.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }

    internal static class XElementStringExtensions
    {
        public static string ElementAttributeValue(this string input, string attribute)
        {
            var element = XElement.Parse(input);
            return element.Attribute(attribute) != null ? element.Attribute(attribute).Value : null;
        }
    }
}