using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration.SyntexModels.Models;
using System;
using System.IO;
using System.Linq;
using System.Web;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSyntexModels: ObjectHandlerBase
    {
        private const string Field_FileLeafRef = "FileLeafRef";

        private const string Field_ModelExplanations = "ModelExplanations";
        private const string Field_ModelDescription = "ModelDescription";
        private const string Field_ModelSchemas = "ModelSchemas";
        private const string Field_ModelMappedClassifierName = "ModelMappedClassifierName";
        private const string Field_ModelLastTrained = "ModelLastTrained";
        private const string Field_ModelSettings = "ModelSettings";
        private const string Field_ModelConfidenceScore = "ModelConfidenceScore";
        private const string Field_ModelAccuracy = "ModelAccuracy";
        private const string Field_ModelClassifiedItemCount = "ModelClassifiedItemCount";
        private const string Field_ModelMismatchedItemCount = "ModelMismatchedItemCount";

        private const string Field_SampleMarkups = "SampleMarkups";
        private const string Field_SampleModelId = "SampleModelId";
        private const string Field_SampleDescription = "SampleDescription";
        private const string Field_SampleExtractedText = "SampleExtractedText";
        private const string Field_SampleFileType = "SampleFileType";
        private const string Field_SampleLabelUptime = "SampleLabelUptime";
        private const string Field_SampleTokenEndPosition = "SampleTokenEndPosition";
        private const string Field_SampleTokenStartPosition = "SampleTokenStartPosition";

        public override string Name
        {
            get { return "SyntexModels"; }
        }

        public override string InternalName => "SyntexModels";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(p => p.Url, p => p.ServerRelativeUrl, p => p.ServerRelativePath);

                var modelLibrary = GetModelLibrary(web);
                var trainingLibrary = GetTrainingLibrary(web);
                if (modelLibrary != null)
                {
                    // Load all models
                    var models = modelLibrary.GetItems(CamlQuery.CreateAllItemsQuery());
                    web.Context.Load(models, p => p.IncludeWithDefaultProperties(p => p.File));
                    web.Context.ExecuteQueryRetry();

                    foreach(var model in models)
                    {
                        if (creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.SyntexModels != null && creationInfo.ExtractConfiguration.SyntexModels.Models != null)
                        {
                            foreach (var modelToExtract in creationInfo.ExtractConfiguration.SyntexModels.Models)
                            {
                                if ((!string.IsNullOrEmpty(modelToExtract.Name) && 
                                    model.FieldValues[Field_FileLeafRef] != null && 
                                    model.FieldValues[Field_FileLeafRef].ToString().Equals(modelToExtract.Name, StringComparison.InvariantCultureIgnoreCase)) ||
                                    (modelToExtract.Id > 0 && model.Id == modelToExtract.Id))
                                {
                                    ExtractModel(web, template, creationInfo, scope, modelToExtract, model, trainingLibrary);
                                }
                            }
                        }
                        else
                        {
                            ExtractModel(web, template, creationInfo, scope, null, model, trainingLibrary);
                        }
                    }
                }
            }

            return template;
        }

        private void ExtractModel(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, ExtractSyntexModelsModelsConfiguration extractConfiguration, ListItem model, List trainingLibrary)
        {
            bool excludeTrainingData = false;
            if (extractConfiguration != null)
            {
                excludeTrainingData = extractConfiguration.ExcludeTrainingData;
            }

            // Add the model file to the template
            (bool fileAdded, string filePath, string fileName) = LoadAndAddSyntexFile(web, model.File, template, creationInfo, scope);

            // Export model metadata
            var templateFile = template.Files.FirstOrDefault(p => p.Src == $"{filePath}/{fileName}");
            if (templateFile != null)
            {
                if (model.FieldValues[Field_ModelExplanations] != null)
                {
                    templateFile.Properties.Add(Field_ModelExplanations, model.FieldValues[Field_ModelExplanations].ToString());
                }
                if (model.FieldValues[Field_ModelSchemas] != null)
                {
                    templateFile.Properties.Add(Field_ModelSchemas, model.FieldValues[Field_ModelSchemas].ToString());
                }
                if (model.FieldValues[Field_ModelDescription] != null)
                {
                    templateFile.Properties.Add(Field_ModelDescription, model.FieldValues[Field_ModelDescription].ToString());
                }
                if (model.FieldValues[Field_ModelMappedClassifierName] != null)
                {
                    templateFile.Properties.Add(Field_ModelMappedClassifierName, model.FieldValues[Field_ModelMappedClassifierName].ToString());
                }
                if (model.FieldValues[Field_ModelLastTrained] != null)
                {
                    templateFile.Properties.Add(Field_ModelLastTrained, model.FieldValues[Field_ModelLastTrained].ToString());
                }
                if (model.FieldValues[Field_ModelSettings] != null)
                {
                    templateFile.Properties.Add(Field_ModelSettings, model.FieldValues[Field_ModelSettings].ToString());
                }
                if (model.FieldValues[Field_ModelConfidenceScore] != null)
                {
                    templateFile.Properties.Add(Field_ModelConfidenceScore, model.FieldValues[Field_ModelConfidenceScore].ToString());
                }
                if (model.FieldValues[Field_ModelAccuracy] != null)
                {
                    templateFile.Properties.Add(Field_ModelAccuracy, model.FieldValues[Field_ModelAccuracy].ToString());
                }
                if (model.FieldValues[Field_ModelClassifiedItemCount] != null)
                {
                    templateFile.Properties.Add(Field_ModelClassifiedItemCount, model.FieldValues[Field_ModelClassifiedItemCount].ToString());
                }
                if (model.FieldValues[Field_ModelMismatchedItemCount] != null)
                {
                    templateFile.Properties.Add(Field_ModelMismatchedItemCount, model.FieldValues[Field_ModelMismatchedItemCount].ToString());
                }
            }

            // Extract training files
            if (trainingLibrary != null && !excludeTrainingData)
            {
                var camlQuery = new CamlQuery
                {
                    ViewXml = string.Format(
                                @"<View Scope='RecursiveAll'>
                                    <Query>
                                        <Where>
                                            <Eq>
                                                <FieldRef Name='SampleModelId' LookupId='TRUE'/>
                                                <Value Type='text'>{0}</Value>
                                            </Eq>
                                        </Where>
                                    </Query>
                                    <ViewFields>
                                        <FieldRef Name='Title'/>
                                        <FieldRef Name='SampleMarkups' />
                                        <FieldRef Name='SampleDescription' />
                                        <FieldRef Name='SampleExtractedText' />
                                        <FieldRef Name='SampleFileType' />
                                        <FieldRef Name='SampleLabelUptime' />
                                        <FieldRef Name='SampleTokenEndPosition' />
                                        <FieldRef Name='SampleTokenStartPosition' />
                                    </ViewFields>
                                  </View>", model.Id)
                };

                // Load training files associated to this model
                var trainingFiles = trainingLibrary.GetItems(camlQuery);
                web.Context.Load(trainingFiles, p => p.IncludeWithDefaultProperties(p => p.File, p=>p.FieldValuesAsText));
                web.Context.ExecuteQueryRetry();

                foreach(var trainingFile in trainingFiles)
                {
                    // Export training files
                    (bool trainingFileAdded, string trainingFilePath, string trainingFileName) = LoadAndAddSyntexFile(web, trainingFile.File, template, creationInfo, scope);
                    var templateTrainingFile = template.Files.FirstOrDefault(p => p.Src == $"{trainingFilePath}/{trainingFileName}");
                    if (templateTrainingFile != null)
                    {
                        // Export training file metadata
                        templateTrainingFile.Properties.Add(Field_SampleModelId, $"{{filelistitemid:{templateFile.Src}}}");
                        if (trainingFile.FieldValues[Field_SampleMarkups] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleMarkups, TokenizeSampleMarkups(trainingFile.FieldValues[Field_SampleMarkups].ToString(), model.Id, $"{{filelistitemid:{templateFile.Src}}}"));
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleDescription) && trainingFile.FieldValues[Field_SampleDescription] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleDescription, trainingFile.FieldValues[Field_SampleDescription].ToString());
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleExtractedText) && trainingFile.FieldValues[Field_SampleExtractedText] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleExtractedText, trainingFile.FieldValues[Field_SampleExtractedText].ToString());
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleFileType) && trainingFile.FieldValues[Field_SampleFileType] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleFileType, trainingFile.FieldValues[Field_SampleFileType].ToString());
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleLabelUptime) && trainingFile.FieldValues[Field_SampleLabelUptime] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleLabelUptime, trainingFile.FieldValues[Field_SampleLabelUptime].ToString());
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleTokenEndPosition) && trainingFile.FieldValues[Field_SampleTokenEndPosition] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleTokenEndPosition, trainingFile.FieldValues[Field_SampleTokenEndPosition].ToString());
                        }
                        if (trainingFile.FieldValues.ContainsKey(Field_SampleTokenStartPosition) && trainingFile.FieldValues[Field_SampleTokenStartPosition] != null)
                        {
                            templateTrainingFile.Properties.Add(Field_SampleTokenStartPosition, trainingFile.FieldValues[Field_SampleTokenStartPosition].ToString());
                        }
                    }
                }
            }
        }

        private string TokenizeSampleMarkups(string sampleMarkupJson, int modelId, string tokenValue)
        {
            sampleMarkupJson = sampleMarkupJson.Replace($"\"{modelId}\": {{", $"\"{tokenValue}\": {{");
            sampleMarkupJson = sampleMarkupJson.Replace($"\"{modelId}\":{{", $"\"{tokenValue}\": {{");
            sampleMarkupJson = sampleMarkupJson.Replace($"\"modelItemId\": \"{modelId}\"", $"\"modelItemId\": \"{tokenValue}\"");
            sampleMarkupJson = sampleMarkupJson.Replace($"\"modelItemId\":{modelId}", $"\"modelItemId\": {tokenValue}");
            return sampleMarkupJson;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            throw new NotImplementedException();
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            web.EnsureProperties(p => p.WebTemplate);

            // Only run extractor when we're in a content center template
            if (web.WebTemplate == "CONTENTCTR")
            {
                return true;
            }

            return false;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            throw new NotImplementedException();
        }

        private List GetModelLibrary(Web web)
        {
            var lists = web.Lists;
            web.Context.Load(lists, list => list.Where(l => l.BaseTemplate == (int)1328).Include(l => l.Id, l => l.Fields));
            web.Context.ExecuteQueryRetry();

            return lists.SingleOrDefault();
        }

        private List GetTrainingLibrary(Web web)
        {
            var lists = web.Lists;
            web.Context.Load(lists, list => list.Where(l => l.BaseTemplate == (int)1330).Include(l => l.Id, l => l.Fields));
            web.Context.ExecuteQueryRetry();

            return lists.SingleOrDefault();
        }

        private Tuple<bool, string, string> LoadAndAddSyntexFile(Web web, Microsoft.SharePoint.Client.File syntexFile, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var baseUri = new Uri(web.Url);

            syntexFile.EnsureProperty(p => p.ServerRelativePath);

            var fullUri = new Uri(baseUri, syntexFile.ServerRelativePath.DecodedUrl);
            var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
            var fileName = Uri.UnescapeDataString(fullUri.Segments[fullUri.Segments.Length - 1]);

            var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

            // Avoid duplicate file entries
            var fileAlreadyExported = template.Files.Where(p => p.Folder.Equals(templateFolderPath, StringComparison.CurrentCultureIgnoreCase) &&
                                                                p.Src.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault();
            if (fileAlreadyExported == null)
            {
                // Add a File to the template
                template.Files.Add(new Model.File()
                {
                    Folder = templateFolderPath,
                    Src = $"{templateFolderPath}/{fileName}",
                    Overwrite = true,
                    Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), syntexFile.Level.ToString())
                });

                // Export the file
                PersistFile(web, creationInfo, scope, folderPath, fileName);

                return new Tuple<bool, string, string>(true, templateFolderPath, fileName);
            }

            return new Tuple<bool, string, string>(false, templateFolderPath, fileName);
        }

        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName)
        {
            if (creationInfo.FileConnector != null)
            {
                var fileConnector = creationInfo.FileConnector;
                SharePointConnector connector = new SharePointConnector(web.Context, web.Url, "dummy");
                Uri u = new Uri(web.Url);

                if (u.PathAndQuery != "/")
                {
                    if (folderPath.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        folderPath = folderPath.Replace(u.PathAndQuery, "");
                    }
                }

                folderPath = Uri.UnescapeDataString(folderPath);
                string container = Uri.UnescapeDataString(folderPath).Trim('/').Replace("/", "\\");
                string persistenceFileName = Uri.UnescapeDataString(fileName);

                if (fileConnector.Parameters.ContainsKey(FileConnectorBase.CONTAINER))
                {
                    container = string.Concat(fileConnector.GetContainer(), container);
                }

                using (Stream s = connector.GetFileStream(persistenceFileName, folderPath))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(
                            persistenceFileName, container, s);
                    }
                }
            }
            else
            {
                scope.LogError($"No connector present to persist {fileName}.");
            }
        }
    }
}
