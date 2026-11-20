using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Entities;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectComposedLook : ObjectHandlerBase
    {
        private const string ComposedLookInfoKey = "_PnP_ProvisioningTemplateComposedLookInfo";
        private static readonly Regex AssetParameter = new Regex(@"\{(?:parameter:|\$)([^{}]+)\}", RegexOptions.IgnoreCase);

        public override string Name
        {
            get { return "Composed Looks"; }
        }

        public override string InternalName => "ComposedLooks";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.ComposedLook != null &&
                    !template.ComposedLook.IsEmptyOrBlank())
                {
                    // Check if this is not a noscript site as themes and composed looks are not supported
                    if (web.IsNoScriptSite())
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_NoSiteCheck);
                        return parser;
                    }

                    bool executeQueryNeeded = false;
                    if (executeQueryNeeded)
                    {
                        web.Context.ExecuteQueryRetry();
                    }

                    (string Palette, string Font, string Background) appliedAssets;
                    if (String.IsNullOrEmpty(template.ComposedLook.ColorFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.FontFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                    {
                        // Apply OOB theme
                        appliedAssets = web.SetComposedLookByUrl(template.ComposedLook.Name, "", "", "", null, false, true, applyingInformation.PropertyBagWriteAllowed);
                    }
                    else
                    {
                        // Apply custom theme
                        string colorFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            colorFile = parser.ParseString(template.ComposedLook.ColorFile);
                        }
                        string backgroundFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            backgroundFile = parser.ParseString(template.ComposedLook.BackgroundFile);
                        }
                        string fontFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            fontFile = parser.ParseString(template.ComposedLook.FontFile);
                        }

                        string masterUrl = null;
                        if (template.WebSettings != null && !string.IsNullOrEmpty(template.WebSettings.MasterPageUrl))
                        {
                            masterUrl = parser.ParseString(template.WebSettings.MasterPageUrl);
                        }
                        web.CreateComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                        appliedAssets = web.SetComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl, false, true, applyingInformation.PropertyBagWriteAllowed);
                    }

                    // Persist composed look info in property bag
                    if (applyingInformation.PropertyBagWriteAllowed != false)
                    {
                        // Keep the original template fields readable by older versions.
                        var metadata = JObject.FromObject(template.ComposedLook);
                        metadata["AppliedAssets"] = JObject.FromObject(new
                        {
                            Theme = appliedAssets.Palette,
                            Font = appliedAssets.Font,
                            BackgroundImage = appliedAssets.Background
                        });
                        var parameters = GetAssetParameters(template.ComposedLook, parser);
                        if (parameters.Count > 0)
                        {
                            metadata["AssetParameters"] = JObject.FromObject(parameters);
                        }
                        web.SetPropertyBagValue(ComposedLookInfoKey, metadata.ToString(Formatting.None));
                    }
                    else
                    {
                        ReportSkippedPropertyBagWrite?.Invoke("composed look properties");
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Retrieving_current_composed_look);

                // Ensure that we have URL property loaded for web and site
                web.EnsureProperty(w => w.Url);
                Site site = (web.Context as ClientContext).Site;
                site.EnsureProperty(s => s.Url);

                SharePointConnector spConnector = new SharePointConnector(web.Context, web.Url, "dummy");
                // to get files from theme catalog we need a connector linked to the root site
                SharePointConnector spConnectorRoot;
                if (!site.Url.Equals(web.Url, StringComparison.InvariantCultureIgnoreCase))
                {
                    spConnectorRoot = new SharePointConnector(web.Context.Clone(site.Url), site.Url, "dummy");
                }
                else
                {
                    spConnectorRoot = spConnector;
                }

                var composedLook = ReadStoredComposedLook(web, scope, out var appliedAssets, out var parameters);
                var currentLook = web.GetCurrentComposedLookForExtraction(site.Url, composedLook?.Name);
                bool matches = composedLook != null && MatchesCurrentLook(web, composedLook, appliedAssets, currentLook, template.Parameters);

                if (currentLook != null)
                {
                    currentLook.Theme = ToAssetPath(currentLook.Theme);
                    currentLook.Font = ToAssetPath(currentLook.Font);
                    currentLook.BackgroundImage = ToAssetPath(currentLook.BackgroundImage);
                }

                if (matches)
                {
                    scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Using_ComposedLookInfoFromPropertyBag);

                    composedLook.BackgroundFile = TokenizeAsset(composedLook.BackgroundFile, web.Url);
                    composedLook.FontFile = TokenizeAsset(composedLook.FontFile, web.Url);
                    composedLook.ColorFile = TokenizeAsset(composedLook.ColorFile, web.Url);
                    template.ComposedLook = composedLook;

                    if (parameters != null)
                    {
                        foreach (var parameter in parameters)
                        {
                            if (!template.Parameters.ContainsKey(parameter.Key))
                            {
                                template.Parameters.Add(parameter.Key, parameter.Value);
                            }
                        }
                    }

                    bool download = creationInfo?.PersistBrandingFiles == true && creationInfo.FileConnector != null && !web.IsSubSite();
                    var files = new[]
                    {
                        (composedLook.BackgroundFile, currentLook.BackgroundImage),
                        (composedLook.ColorFile, currentLook.Theme),
                        (composedLook.FontFile, currentLook.Font)
                    };
                    foreach (var (expression, path) in files)
                    {
                        if (!string.IsNullOrEmpty(expression) && !string.IsNullOrEmpty(path))
                        {
                            var file = GetComposedLookFile(path);
                            file.Folder = Tokenize(file.Folder, web.Url);
                            template.Files.Add(file);
                            if (download)
                            {
                                DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, path, scope);
                            }
                        }
                    }
                }
                else
                {
                    template = ExtractCurrentComposedLook(web, template, creationInfo, scope, spConnector, spConnectorRoot, currentLook);
                }

                if (creationInfo != null && creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private string TokenizeAsset(string asset, string webUrl)
        {
            return asset?.IndexOf('{') >= 0 ? asset : Tokenize(ToAssetPath(asset), webUrl);
        }

        private static string ToAssetPath(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
                (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp))
            {
                return Uri.UnescapeDataString(uri.AbsolutePath);
            }

            return url;
        }

        private static Dictionary<string, string> GetAssetParameters(ComposedLook look, TokenParser parser)
        {
            var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var asset in new[] { look.ColorFile, look.FontFile, look.BackgroundFile })
            {
                if (string.IsNullOrEmpty(asset))
                {
                    continue;
                }
                foreach (Match token in AssetParameter.Matches(asset))
                {
                    var value = parser.ParseString(token.Value);
                    if (value != token.Value)
                    {
                        parameters[token.Groups[1].Value] = value;
                    }
                }
            }
            return parameters;
        }

        private static ComposedLook ReadStoredComposedLook(Web web, PnPMonitoredScope scope, out ThemeEntity appliedAssets, out Dictionary<string, string> parameters)
        {
            appliedAssets = null;
            parameters = null;
            var json = web.GetPropertyBagValueString(ComposedLookInfoKey, null);
            if (string.IsNullOrEmpty(json))
            {
                return null;
            }

            try
            {
                var metadata = JObject.Parse(json);
                var look = metadata.ToObject<ComposedLook>();
                if (look?.Name != null)
                {
                    appliedAssets = metadata["AppliedAssets"]?.ToObject<ThemeEntity>();
                    parameters = metadata["AssetParameters"]?.ToObject<Dictionary<string, string>>();
                    return look;
                }
            }
            catch (JsonException)
            {
                // Fall back to the active gallery entry if the saved metadata cannot be read.
            }

            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_ComposedLookInfoFailedToDeserialize);
            return null;
        }

        private static bool MatchesCurrentLook(Web web, ComposedLook stored, ThemeEntity appliedAssets, ThemeEntity current,
            Dictionary<string, string> templateParameters)
        {
            if (current == null || !string.Equals(stored.Name, current.Name, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (appliedAssets == null)
            {
                // Older entries contain only template fields. Built-in looks can be identified by name.
                if (!current.IsCustomComposedLook && string.IsNullOrEmpty(stored.ColorFile) &&
                    string.IsNullOrEmpty(stored.FontFile) && string.IsNullOrEmpty(stored.BackgroundFile))
                {
                    return true;
                }
                var legacyTemplate = new ProvisioningTemplate { ComposedLook = stored };
                foreach (var parameter in templateParameters)
                {
                    legacyTemplate.Parameters.Add(parameter.Key, parameter.Value);
                }
                var parser = new TokenParser(web, legacyTemplate);
                appliedAssets = new ThemeEntity
                {
                    Theme = parser.ParseString(stored.ColorFile),
                    Font = parser.ParseString(stored.FontFile),
                    BackgroundImage = parser.ParseString(stored.BackgroundFile)
                };
            }

            var baseUrl = new Uri(web.Url.TrimEnd('/') + "/");
            return MatchesAsset(baseUrl, appliedAssets.Theme, current.Theme) &&
                MatchesAsset(baseUrl, appliedAssets.Font, current.Font) &&
                MatchesAsset(baseUrl, appliedAssets.BackgroundImage, current.BackgroundImage);
        }

        private static bool MatchesAsset(Uri baseUrl, string expectedUrl, string currentUrl)
        {
            if (string.IsNullOrEmpty(expectedUrl) || string.IsNullOrEmpty(currentUrl))
            {
                return string.IsNullOrEmpty(expectedUrl) && string.IsNullOrEmpty(currentUrl);
            }

            return Uri.TryCreate(baseUrl, expectedUrl, out var storedAsset) &&
                Uri.TryCreate(baseUrl, currentUrl, out var currentAsset) &&
                string.Equals(Uri.UnescapeDataString(storedAsset.GetLeftPart(UriPartial.Path)),
                    Uri.UnescapeDataString(currentAsset.GetLeftPart(UriPartial.Path)), StringComparison.OrdinalIgnoreCase);
        }

        private ProvisioningTemplate ExtractCurrentComposedLook(Web web, ProvisioningTemplate template,
                                                    ProvisioningTemplateCreationInformation creationInfo,
                                                    PnPMonitoredScope scope, SharePointConnector spConnector,
                                                    SharePointConnector spConnectorRoot, ThemeEntity theme)
        {
            if (theme != null)
            {
                if (creationInfo != null)
                {
                    // Don't exclude the DesignPreviewThemedCssFolderUrl property bag, if any
                    creationInfo.PropertyBagPropertiesToPreserve.Add("DesignPreviewThemedCssFolderUrl");
                }

                template.ComposedLook.Name =
                    theme.Name != null ? theme.Name : String.Empty;

                if (theme.IsCustomComposedLook)
                {
                    // Set the URL pointers to files
                    template.ComposedLook.BackgroundFile = FixFileUrl(Tokenize(theme.BackgroundImage, web.Url));
                    template.ComposedLook.ColorFile = FixFileUrl(Tokenize(theme.Theme, web.Url));
                    template.ComposedLook.FontFile = FixFileUrl(Tokenize(theme.Font, web.Url));

                    // Download files if this is root site, since theme files are only stored there
                    if (!web.IsSubSite() && creationInfo != null &&
                        creationInfo.PersistBrandingFiles && creationInfo.FileConnector != null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                        // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                        // Download the theme/branding specific files
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.BackgroundImage, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Theme, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web.Url, theme.Font, scope);
                    }

                    // Create file entries for the custom theme files, but only if it's a root site
                    // If it's root site we do not extract or set theme files, since those are in the root of the site collection
                    if (!web.IsSubSite())
                    {
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.BackgroundFile));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.ColorFile));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.FontFile));
                        }
                    }
                    // If a base template is specified then use that one to "cleanup" the generated template model
                    if (creationInfo != null && creationInfo.BaseTemplate != null)
                    {
                        template = CleanupEntities(template, creationInfo.BaseTemplate);
                    }
                }
                else
                {
                    template.ComposedLook.BackgroundFile = "";
                    template.ComposedLook.ColorFile = "";
                    template.ComposedLook.FontFile = "";
                }
            }
            else
            {
                template.ComposedLook = null;
            }

            return template;
        }

        private void DownLoadFile(SharePointConnector reader, SharePointConnector readerRoot, FileConnectorBase writer, string webUrl, string asset, PnPMonitoredScope scope)
        {

            // No file passed...leave
            if (String.IsNullOrEmpty(asset))
            {
                return;
            }

            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_DownLoadFile_Downloading_asset___0_, asset);
            ;

            SharePointConnector readerToUse;
            Model.File f = GetComposedLookFile(asset);

            // Strip the /sites/root part from /sites/root/lib/folder structure, special case for root site handling.
            Uri u = new Uri(webUrl);
            if (f.Folder.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1 && u.PathAndQuery.Length > 1)
            {
                f.Folder = f.Folder.Replace(u.PathAndQuery, "");
            }

            // in case of a theme catalog we need to use the root site reader as that list only exists on root site level
            if (f.Folder.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                readerToUse = readerRoot;
            }
            else
            {
                readerToUse = reader;
            }

            using (Stream s = readerToUse.GetFileStream(f.Src, f.Folder))
            {
                if (s != null)
                {
                    writer.SaveFileStream(f.Src, s);
                }
            }
        }

        private static string FixFileName(string originalFileName)
        {
            // if we've found the file use the provided writer to persist the downloaded file
            String regexStrip = @"(\\|/|:|\*|\?|""|>|<|\||=)*";
            String result = Regex.Replace(originalFileName.Substring(0,
                originalFileName.IndexOf("?") > 0 ? originalFileName.IndexOf("?") : originalFileName.Length),
                regexStrip, "", RegexOptions.IgnorePatternWhitespace);

            return result;
        }
        private string FixFileUrl(string originalFileUrl)
        {
            if (string.IsNullOrEmpty(originalFileUrl))
            {
                return "";
            }

            String fileUrl = originalFileUrl.Substring(0, originalFileUrl.LastIndexOf("/"));
            String fileName = FixFileName(originalFileUrl.Substring(originalFileUrl.LastIndexOf("/") + 1));

            String result = $"{fileUrl}/{fileName}";

            return result;
        }

        private Model.File GetComposedLookFile(string asset)
        {
            int index = asset.LastIndexOf("/");
            Model.File file = new Model.File
            {
                Src = FixFileName(asset.Substring(index + 1)),
                Folder = asset.Substring(0, index),
                Overwrite = true,
                Security = null
            };
            return file;
        }

        private static ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            if (template.ComposedLook != null && baseTemplate.ComposedLook != null)
            {
                if (template.ComposedLook.Equals(baseTemplate.ComposedLook))
                {
                    template.ComposedLook = null;
                }

            }
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = (template.ComposedLook != null && !template.ComposedLook.IsEmptyOrBlank() && !web.IsNoScriptSite());
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = !web.IsNoScriptSite();
            }
            return _willExtract.Value;
        }
    }
}
