using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization.Telemetry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace PnP.Framework.Modernization.Functions
{
    public partial class BuiltIn
    {
        /// <summary>
        /// Maps Project Summary to the destination's live task lists and configured timeline.
        /// </summary>
        [FunctionDocumentation(Description = "Maps Project Summary to target task lists. Timeline membership is captured at migration; task values remain live.",
            Example = "{JsonProperties} = ProjectSummaryProperties({ListId}, {PrimaryTaskListUrl}, {ShowTimeline}, {Panels}, {Title})")]
        public string ProjectSummaryProperties(Guid listId, string primaryTaskListUrl, bool showTimeline, string panels, string title)
        {
            var source = sourceClientContext ?? clientContext;
            List sourceList;
            if (listId != Guid.Empty)
            {
                sourceList = source.Web.GetListById(listId);
                if (sourceList == null && !string.IsNullOrWhiteSpace(primaryTaskListUrl))
                {
                    sourceList = ProjectSummarySourceList(primaryTaskListUrl);
                }
            }
            else if (!string.IsNullOrWhiteSpace(primaryTaskListUrl))
            {
                sourceList = ProjectSummarySourceList(primaryTaskListUrl);
            }
            else
            {
                var lists = source.LoadQuery(source.Web.Lists.Where(list =>
                    list.BaseTemplate == (int)ListTemplateType.TasksWithTimelineAndHierarchy));
                source.ExecuteQueryRetry();
                var candidates = lists.ToArray();
                if (candidates.Length != 1)
                {
                    throw new NotAvailableAtTargetException("Project Summary has no explicit primary task list and cannot resolve a unique default task list. This web part will be skipped.");
                }
                sourceList = candidates[0];
            }

            var targetList = ProjectSummaryTargetList(sourceList);
            targetList.EnsureProperties(list => list.Id, list => list.BaseTemplate, list => list.RootFolder);
            if (targetList.BaseTemplate != (int)ListTemplateType.Tasks &&
                targetList.BaseTemplate != (int)ListTemplateType.TasksWithTimelineAndHierarchy)
            {
                throw new NotAvailableAtTargetException("Project Summary requires a task list at the destination. This web part will be skipped.");
            }
            targetList.RootFolder.EnsureProperty(folder => folder.ServerRelativeUrl);
            var properties = new JObject
            {
                ["title"] = title ?? string.Empty,
                ["listId"] = targetList.Id.ToString(),
                ["primaryTaskListUrl"] = targetList.RootFolder.ServerRelativeUrl,
                ["showTimeline"] = showTimeline,
                ["panels"] = ProjectSummaryPanels(panels)
            };

            if (showTimeline)
            {
                targetList.RootFolder.EnsureProperty(folder => folder.Properties);
                AddProjectSummaryTimeline(properties, targetList.RootFolder.Properties.FieldValues);
                LogWarning("Project Summary uses the destination task list's timeline membership and date range captured at migration. Task values reload live, but later timeline membership or configuration changes require remigration. Classic timeline styling is not supported.",
                    LogStrings.Heading_BuiltInFunctions);
            }

            return new JObject
            {
                ["dataVersion"] = "1.0",
                ["properties"] = properties
            }.ToString(Newtonsoft.Json.Formatting.None);
        }

        private List ProjectSummarySourceList(string listUrl)
        {
            var source = sourceClientContext ?? clientContext;
            var webUrl = new Uri(source.Url.TrimEnd('/') + "/");
            var resolved = new Uri(webUrl, listUrl);
            if (!webUrl.IsBaseOf(resolved))
            {
                throw new NotAvailableAtTargetException($"Project Summary list URL '{listUrl}' is outside the source web. This web part will be skipped.");
            }
            return source.Web.GetList(Uri.UnescapeDataString(resolved.AbsolutePath));
        }

        private List ProjectSummaryTargetList(List sourceList)
        {
            if (sourceList == null)
            {
                throw new NotAvailableAtTargetException("Project Summary's source task list could not be found. This web part will be skipped.");
            }
            sourceList.EnsureProperty(list => list.Id);
            var source = sourceClientContext ?? clientContext;
            if (source.Url.TrimEnd('/').Equals(clientContext.Url.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
            {
                return sourceList;
            }
            return clientContext.Web.GetListById(Guid.Parse(ListCrossSiteCheck(sourceList.Id)));
        }

        private JArray ProjectSummaryPanels(string panels)
        {
            if (string.IsNullOrWhiteSpace(panels))
            {
                throw new NotAvailableAtTargetException("Project Summary panel settings could not be read. Enable web part export before migration. This web part will be skipped.");
            }

            var result = new JArray();
            foreach (var item in JArray.Parse(panels))
            {
                var panel = item as JObject ?? throw new FormatException("Project Summary panel settings must contain objects.");
                var urls = new JArray();
                var listUrls = panel.GetValue("ListUrls", StringComparison.OrdinalIgnoreCase);
                if (listUrls != null && listUrls.Type != JTokenType.Null)
                {
                    if (!(listUrls is JArray listArray))
                    {
                        throw new FormatException("Project Summary ListUrls must be an array.");
                    }
                    foreach (var url in listArray)
                    {
                        if (url.Type != JTokenType.String || string.IsNullOrWhiteSpace((string)url))
                        {
                            throw new FormatException("Project Summary panel list URLs must be nonempty strings.");
                        }
                        var target = ProjectSummaryTargetList(ProjectSummarySourceList((string)url));
                        target.EnsureProperty(list => list.RootFolder);
                        target.RootFolder.EnsureProperty(folder => folder.ServerRelativeUrl);
                        urls.Add(target.RootFolder.ServerRelativeUrl);
                    }
                }
                var enabled = ProjectSummaryPanelValue(panel, "Enabled", JTokenType.Boolean);
                var late = ProjectSummaryPanelValue(panel, "IsLatePanel", JTokenType.Boolean);
                var days = ProjectSummaryPanelValue(panel, "DaysToSearch", JTokenType.Integer);
                result.Add(new JObject
                {
                    ["enabled"] = enabled,
                    ["title"] = (string)panel.GetValue("Title", StringComparison.OrdinalIgnoreCase) ?? string.Empty,
                    ["isLatePanel"] = late,
                    ["daysToSearch"] = days,
                    ["listUrls"] = urls
                });
            }

            // The native component replaces an empty array with enabled default panels.
            if (result.Count == 0)
            {
                result.Add(new JObject
                {
                    ["enabled"] = false, ["title"] = string.Empty, ["isLatePanel"] = true,
                    ["daysToSearch"] = 0, ["listUrls"] = new JArray()
                });
            }
            return result;
        }

        private static JToken ProjectSummaryPanelValue(JObject panel, string name, JTokenType type)
        {
            var value = panel.GetValue(name, StringComparison.OrdinalIgnoreCase);
            if (value == null || value.Type != type)
            {
                throw new FormatException($"Project Summary panel {name} must be a {type}.");
            }
            return value.DeepClone();
        }

        internal static void AddProjectSummaryTimeline(JObject properties, IDictionary<string, object> folderProperties)
        {
            string view = folderProperties.TryGetValue("TimelineSelectedView", out var selected) ? selected as string : null;
            if (string.IsNullOrEmpty(view))
            {
                view = folderProperties.TryGetValue("TimelineDefaultView", out var defaultView) ? defaultView as string : null;
            }
            string key = "Timeline_" + view;
            if (!folderProperties.ContainsKey(key))
            {
                var configured = folderProperties.Keys.Where(name => name.StartsWith("Timeline_", StringComparison.Ordinal)).ToArray();
                if (configured.Length > 1)
                {
                    throw new NotAvailableAtTargetException("Project Summary cannot identify the destination's selected timeline view. This web part will be skipped.");
                }
                key = configured.SingleOrDefault();
                view = key?.Substring("Timeline_".Length) ?? string.Empty;
            }

            var ids = new JArray();
            properties["timelineViewName"] = view;
            properties["timelineItemIds"] = ids;
            if (key == null)
            {
                return;
            }
            var xml = folderProperties[key] as string;
            if (string.IsNullOrWhiteSpace(xml))
            {
                throw new FormatException("Project Summary timeline configuration is empty or unreadable.");
            }
            XElement timeline;
            using (var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
            {
                timeline = XElement.Load(reader);
            }
            if (timeline.Name.LocalName != "TLViewData")
            {
                throw new FormatException("Project Summary timeline configuration must contain TLViewData.");
            }
            var members = new HashSet<int>();
            foreach (var item in timeline.Descendants().Where(element =>
                (element.Name.LocalName == "t" || element.Name.LocalName == "ft" || element.Name.LocalName == "m") &&
                (string)element.Attribute("onTL") == "1"))
            {
                var value = (string)item.Attribute("uid");
                if (value == "0")
                {
                    continue;
                }
                if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id <= 0)
                {
                    throw new FormatException("Project Summary timeline membership contains an invalid task ID.");
                }
                if (members.Add(id))
                {
                    ids.Add(id);
                }
            }
            var bars = timeline.Descendants("tl").ToArray();
            if (bars.Length > 1)
            {
                throw new NotAvailableAtTargetException("Project Summary cannot preserve multiple timeline bars in the native destination component. This web part will be skipped.");
            }
            if (bars.Length == 1 && (string)bars[0].Attribute("useCustomDates") == "1")
            {
                var start = ProjectSummaryTimelineDate(bars[0], "startDate");
                var end = ProjectSummaryTimelineDate(bars[0], "finishDate");
                if (end < start)
                {
                    throw new FormatException("Project Summary timeline finish date precedes its start date.");
                }
                properties["timelineStartDate"] = start.ToString("yyyy-MM-dd'T'00:00:00", CultureInfo.InvariantCulture);
                properties["timelineEndDate"] = end.ToString("yyyy-MM-dd'T'00:00:00", CultureInfo.InvariantCulture);
            }
        }

        private static DateTime ProjectSummaryTimelineDate(XElement bar, string name)
        {
            if (!DateTime.TryParse((string)bar.Attribute(name), CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            {
                throw new FormatException($"Project Summary timeline {name} is invalid.");
            }
            return date.Date;
        }
    }
}
