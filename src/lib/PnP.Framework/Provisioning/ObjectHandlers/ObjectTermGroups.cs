using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectTermGroups : ObjectHandlerBase
    {
        private List<TermGroupHelper.ReusedTerm> reusedTerms;

        public override string Name => "Term Groups";

        public override string InternalName => "TermGroups";
        public override TokenParser ProvisionObjects(Web web, Model.ProvisioningTemplate template, TokenParser parser,
            ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                this.reusedTerms = new List<TermGroupHelper.ReusedTerm>();

                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(web.Context);
                TermStore termStore;
                TermGroup siteCollectionTermGroup;
                IEnumerable<TermGroup> termGroups;

                try
                {
                    termStore = taxSession.GetDefaultKeywordsTermStore();

                    web.Context.Load(termStore, ts => ts.Languages, ts => ts.DefaultLanguage);
                    siteCollectionTermGroup = termStore.GetSiteCollectionGroup((web.Context as ClientContext).Site, false);

                    if (applyingInformation.LoadSiteCollectionTermGroups)
                    {
                        termGroups = termStore.Groups;

                        web.Context.Load(
                            termStore.Groups,
                            groups => groups.Include(
                                group => group.Name,
                                group => group.Id,
                                group => group.TermSets.Include(
                                    termSet => termSet.Name,
                                    termSet => termSet.Id)
                            ));
                    }
                    else
                    {
                        termGroups = web.Context.LoadQuery(termStore
                            .Groups
                            .Where(group => !group.IsSiteCollectionGroup)
                            .Include(
                                group => group.Name,
                                group => group.Id,
                                group => group.TermSets.Include(
                                    termSet => termSet.Name,
                                    termSet => termSet.Id)));
                    }

                    web.Context.Load(siteCollectionTermGroup);
                    web.Context.ExecuteQueryRetry();
                }
                catch (ServerException)
                {
                    // If the GetDefaultSiteCollectionTermStore method call fails ... raise a specific Warning
                    WriteMessage(CoreResources.Provisioning_ObjectHandlers_TermGroups_Wrong_Configuration, ProvisioningMessageType.Warning);

                    // and exit skipping the current handler
                    return parser;
                }

                foreach (var modelTermGroup in template.TermGroups)
                {
                    this.reusedTerms.AddRange(TermGroupHelper.ProcessGroup(web.Context as ClientContext, taxSession, termStore, termGroups, modelTermGroup, siteCollectionTermGroup, parser, scope));
                }

                foreach (var reusedTerm in this.reusedTerms)
                {
                    TermGroupHelper.TryReuseTerm(web.Context as ClientContext, reusedTerm.ModelTerm, reusedTerm.Parent, reusedTerm.TermStore, parser, scope);
                }
            }
            return parser;
        }

        private class TryReuseTermResult
        {
            public bool Success { get; set; }
            public TokenParser UpdatedParser { get; set; }
        }

        public override Model.ProvisioningTemplate ExtractObjects(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (creationInfo.IncludeSiteCollectionTermGroup || creationInfo.IncludeAllTermGroups)
                {
                    // Find the site collection termgroup, if any
                    TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);

                    TermStore termStore = null;

                    try
                    {
                        // Refresh / Update the taxonomy Cache to prevent errors using the same session for creating and getting terms
                        session.UpdateCache();
                        web.Context.ExecuteQueryRetry();

                        termStore = session.GetDefaultSiteCollectionTermStore();
                        web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage, t => t.OrphanedTermsTermSet);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (ServerException)
                    {
                        // Skip the exception and go to the next check
                    }

                    if (null == termStore || termStore.ServerObjectIsNull())
                    {
                        // If the GetDefaultSiteCollectionTermStore method call fails ... raise a specific Warning
                        WriteMessage(CoreResources.Provisioning_ObjectHandlers_TermGroups_Wrong_Configuration, ProvisioningMessageType.Warning);

                        // and exit skipping the current handler
                        return template;
                    }

                    var orphanedTermsTermSetId = default(Guid);
                    if (!termStore.OrphanedTermsTermSet.ServerObjectIsNull())
                    {
                        termStore.OrphanedTermsTermSet.EnsureProperty(ts => ts.Id);
                        orphanedTermsTermSetId = termStore.OrphanedTermsTermSet.Id;
                        if (termStore.ServerObjectIsNull.Value)
                        {
                            termStore = session.GetDefaultKeywordsTermStore();
                            web.Context.Load(termStore, t => t.Id, t => t.DefaultLanguage);
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    var propertyBagKey = $"SiteCollectionGroupId{termStore.Id}";

                    // Ensure to grab the property from the rootweb
                    var site = (web.Context as ClientContext).Site;
                    web.Context.Load(site, s => s.RootWeb);
                    web.Context.ExecuteQueryRetry();

                    var siteCollectionTermGroupId = site.RootWeb.GetPropertyBagValueString(propertyBagKey, "");

                    Guid termGroupGuid;
                    Guid.TryParse(siteCollectionTermGroupId, out termGroupGuid);

                    List<TermGroup> termGroups = new List<TermGroup>();
                    if (creationInfo.IncludeAllTermGroups)
                    {
                        web.Context.Load(termStore.Groups, groups => groups.Include(tg => tg.Name,
                            tg => tg.Id,
                            tg => tg.Description,
                            tg => tg.TermSets.IncludeWithDefaultProperties(ts => ts.CustomSortOrder)));
                        web.Context.ExecuteQueryRetry();
                        termGroups = termStore.Groups.ToList();
                    }
                    else
                    {
                        if (termGroupGuid != Guid.Empty)
                        {
                            var termGroup = termStore.GetGroup(termGroupGuid);
                            web.Context.Load(termGroup,
                                tg => tg.Name,
                                tg => tg.Id,
                                tg => tg.Description,
                                tg => tg.TermSets.IncludeWithDefaultProperties(ts => ts.Description, ts => ts.CustomSortOrder));

                            web.Context.ExecuteQueryRetry();

                            termGroups = new List<TermGroup>() { termGroup };
                        }
                    }

                    foreach (var termGroup in termGroups)
                    {
                        Boolean isSiteCollectionTermGroup = termGroupGuid != Guid.Empty && termGroup.Id == termGroupGuid;

                        var modelTermGroup = new Model.TermGroup
                        {
                            Name = isSiteCollectionTermGroup ? "{sitecollectiontermgroupname}" : termGroup.Name,
                            Id = isSiteCollectionTermGroup ? Guid.Empty : termGroup.Id,
                            Description = termGroup.Description,
                            SiteCollectionTermGroup = isSiteCollectionTermGroup
                        };

                        // If we need to include TermGroups security
                        if (creationInfo.IncludeTermGroupsSecurity)
                        {
                            termGroup.EnsureProperties(tg => tg.ContributorPrincipalNames, tg => tg.GroupManagerPrincipalNames);

                            // Extract the TermGroup contributors
                            modelTermGroup.Contributors.AddRange(
                                from c in termGroup.ContributorPrincipalNames
                                select new Model.User { Name = c });

                            // Extract the TermGroup managers
                            modelTermGroup.Managers.AddRange(
                                from m in termGroup.GroupManagerPrincipalNames
                                select new Model.User { Name = m });
                        }

                        web.EnsureProperty(w => w.Url);

                        foreach (var termSet in termGroup.TermSets)
                        {
                            // Do not include the orphan term set
                            if (termSet.Id == orphanedTermsTermSetId) continue;

                            // Extract all other term sets
                            var modelTermSet = new Model.TermSet
                            {
                                Name = termSet.Name
                            };
                            if (!isSiteCollectionTermGroup)
                            {
                                modelTermSet.Id = termSet.Id;
                            }
                            modelTermSet.IsAvailableForTagging = termSet.IsAvailableForTagging;
                            modelTermSet.IsOpenForTermCreation = termSet.IsOpenForTermCreation;
                            modelTermSet.Description = termSet.Description;
                            modelTermSet.Terms.AddRange(GetTerms<TermSet>(web.Context, termSet, termStore.DefaultLanguage, isSiteCollectionTermGroup));
                            foreach (var property in termSet.CustomProperties)
                            {
                                if (property.Key.Equals("_Sys_Nav_AttachedWeb_SiteId", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    modelTermSet.Properties.Add(property.Key, "{sitecollectionid}");
                                }
                                else if (property.Key.Equals("_Sys_Nav_AttachedWeb_WebId", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    modelTermSet.Properties.Add(property.Key, "{siteid}");
                                }
                                else
                                {
                                    modelTermSet.Properties.Add(property.Key, Tokenize(property.Value, web.Url, web));
                                }
                            }
                            modelTermGroup.TermSets.Add(modelTermSet);
                        }

                        template.TermGroups.Add(modelTermGroup);
                    }
                }
            }
            return template;
        }

        private List<Model.Term> GetTerms<T>(ClientRuntimeContext context, TaxonomyItem parent, int defaultLanguage, Boolean isSiteCollectionTermGroup = false)
        {
            List<Model.Term> termsToReturn = new List<Model.Term>();
            TermCollection terms;
            var customSortOrder = string.Empty;
            if (parent is TermSet)
            {
                terms = ((TermSet)parent).Terms;
                customSortOrder = ((TermSet)parent).CustomSortOrder;
            }
            else
            {
                terms = ((Term)parent).Terms;
                customSortOrder = ((Term)parent).CustomSortOrder;
            }
            context.Load(terms, tms => tms.IncludeWithDefaultProperties(t => t.Labels, t => t.CustomSortOrder,
                t => t.IsReused, t => t.IsSourceTerm, t => t.SourceTerm, t => t.IsDeprecated, t => t.Description, t => t.Owner));
            context.ExecuteQueryRetry();

            foreach (var term in terms)
            {
                var modelTerm = new Model.Term();
                if (!isSiteCollectionTermGroup || term.IsReused)
                {
                    modelTerm.Id = term.Id;
                }
                modelTerm.Name = term.Name;
                modelTerm.IsAvailableForTagging = term.IsAvailableForTagging;
                modelTerm.IsReused = term.IsReused;
                modelTerm.IsSourceTerm = term.IsSourceTerm;
                modelTerm.SourceTermId = (term.SourceTerm != null) ? term.SourceTerm.Id : Guid.Empty;
                modelTerm.IsDeprecated = term.IsDeprecated;
                modelTerm.Description = term.Description;
                modelTerm.Owner = term.Owner;

                if ((!term.IsReused || term.IsSourceTerm) && term.Labels.Any())
                {
                    foreach (var label in term.Labels)
                    {
                        if ((label.Language == defaultLanguage && label.Value != term.Name) || label.Language != defaultLanguage)
                        {
                            var modelLabel = new Model.TermLabel
                            {
                                IsDefaultForLanguage = label.IsDefaultForLanguage,
                                Value = label.Value,
                                Language = label.Language
                            };

                            modelTerm.Labels.Add(modelLabel);
                        }
                    }
                }

                foreach (var localProperty in term.LocalCustomProperties)
                {
                    modelTerm.LocalProperties.Add(localProperty.Key, localProperty.Value);
                }

                // Shared Properties have to be extracted just for source terms or not reused terms
                if (!term.IsReused || term.IsSourceTerm)
                {
                    foreach (var customProperty in term.CustomProperties)
                    {
                        modelTerm.Properties.Add(customProperty.Key, customProperty.Value);
                    }
                }

                if (term.TermsCount > 0)
                {
                    modelTerm.Terms.AddRange(GetTerms<Term>(context, term, defaultLanguage, isSiteCollectionTermGroup));
                }
                termsToReturn.Add(modelTerm);

                if (!string.IsNullOrEmpty(customSortOrder))
                {
                    var sortOrder = customSortOrder.Split(new[] { ':' }).ToList();

                    var currentTermIndex = sortOrder.Where(i => new Guid(i) == term.Id).FirstOrDefault();
                    modelTerm.CustomSortOrder = sortOrder.IndexOf(currentTermIndex) + 1;

                }
            }
            termsToReturn = termsToReturn.OrderBy(t => t.CustomSortOrder).ToList();

            return termsToReturn;
        }


        public override bool WillProvision(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.TermGroups.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = creationInfo.IncludeSiteCollectionTermGroup || creationInfo.IncludeAllTermGroups;
            }
            return _willExtract.Value;
        }
    }
}
