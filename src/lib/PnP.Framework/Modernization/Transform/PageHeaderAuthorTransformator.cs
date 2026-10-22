using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Extensions;
using PnP.Framework.Modernization.Telemetry;
using System;
using System.Collections.Generic;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Modernization.Transform
{
    internal sealed class PageHeaderAuthorTransformator : BaseTransform
    {
        private readonly ClientContext sourceContext;
        private readonly ClientContext targetContext;
        private readonly UserTransformator userTransformator;

        internal PageHeaderAuthorTransformator(ClientContext sourceContext, ClientContext targetContext,
            UserTransformator userTransformator, IList<ILogObserver> logObservers)
        {
            this.sourceContext = sourceContext;
            this.targetContext = targetContext;
            this.userTransformator = userTransformator;
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    RegisterObserver(observer);
                }
            }
        }

        internal void SetSourceAuthor(PnPCore.IPageHeader header, ListItem sourcePage)
        {
            if (sourcePage == null ||
                !sourcePage.FieldValues.TryGetValue(Constants.CreatedByField, out var value) ||
                !(value is FieldUserValue sourceAuthor))
            {
                Warn("The source page has no Created By user.");
                return;
            }

            try
            {
                // Resolve by ID, not display name, which need not uniquely identify a person.
                var author = CacheManager.Instance.GetUserFromUserList(sourceContext, sourceAuthor.LookupId);
                if (author == null || author.IsGroup || string.IsNullOrWhiteSpace(author.LoginName))
                {
                    Warn("The source Created By user could not be resolved to a person.");
                    return;
                }

                var principal = userTransformator.RemapPrincipal(author.LoginName);
                if (!TryResolveTargetAuthor(principal, out var resolvedUser, out var targetAuthor))
                {
                    return;
                }

                header.Authors = JsonConvert.SerializeObject(targetAuthor,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                SetByline(header, resolvedUser);
            }
            catch (ServerException ex)
            {
                Warn(ex.Message);
            }
            catch (ClientRequestException ex)
            {
                Warn(ex.Message);
            }
        }

        internal void SetMappedAuthors(PnPCore.IPageHeader header, string authors)
        {
            // Custom mappings own the display payload, including any additional properties.
            header.Authors = authors;
            header.AuthorByLine = null;
            header.AuthorByLineId = -1;
            if (string.IsNullOrWhiteSpace(authors))
            {
                return;
            }

            try
            {
                var token = JToken.Parse(authors);
                if (token is JArray array)
                {
                    if (array.Count == 0)
                    {
                        return;
                    }
                    if (array.Count != 1)
                    {
                        Warn("Multiple mapped authors were preserved, but the single-person page byline was not set.");
                        return;
                    }
                    token = array[0];
                }

                if (!(token is JObject author))
                {
                    Warn("The mapped author must be a JSON object or a single-person array.");
                    return;
                }

                var principal = GetIdentity(author, "upn") ?? GetIdentity(author, "email") ?? GetIdentity(author, "id");
                if (TryResolveTargetAuthor(principal, out var resolvedUser, out _))
                {
                    SetByline(header, resolvedUser);
                }
            }
            catch (JsonException ex)
            {
                Warn($"The mapped author JSON could not be used for the page byline: {ex.Message}");
            }
            catch (ServerException ex)
            {
                Warn(ex.Message);
            }
            catch (ClientRequestException ex)
            {
                Warn(ex.Message);
            }
        }

        private bool TryResolveTargetAuthor(string principal, out ResolvedUser resolvedUser, out UserEntity author)
        {
            resolvedUser = null;
            author = null;
            if (string.IsNullOrWhiteSpace(principal))
            {
                Warn("No user identity was provided for the page byline.");
                return false;
            }

            resolvedUser = CacheManager.Instance.GetEnsuredUser(targetContext, principal);
            if (resolvedUser == null || resolvedUser.Id <= 0 || string.IsNullOrWhiteSpace(resolvedUser.LoginName))
            {
                Warn("The page author could not be resolved in the target site.");
                return false;
            }

            author = CacheManager.Instance.GetUserFromUserList(targetContext, resolvedUser.Id);
            if (author == null || author.IsGroup)
            {
                Warn("The target page author could not be resolved to a person.");
                return false;
            }

            if (!resolvedUser.LoginName.StartsWith("i:0#.f|membership|", StringComparison.OrdinalIgnoreCase) ||
                !resolvedUser.LoginName.GetUserName().Contains("@"))
            {
                Warn("The target identity is not a membership user for the page byline (for example, a system or application account).");
                return false;
            }

            return true;
        }

        private static string GetIdentity(JObject author, string property)
        {
            var value = author[property];
            return value?.Type == JTokenType.String && !string.IsNullOrWhiteSpace((string)value)
                ? ((string)value).Trim()
                : null;
        }

        private static void SetByline(PnPCore.IPageHeader header, ResolvedUser user)
        {
            // PnP Core resolves the byline again on save; pass the user name, not the claims prefix.
            header.AuthorByLine = JsonConvert.SerializeObject(new[] { user.LoginName.GetUserName() });
            header.AuthorByLineId = user.Id;
        }

        private void Warn(string message)
        {
            LogWarning(string.Format(LogStrings.Warning_PageHeaderAuthorNotSet, message),
                LogStrings.Heading_ArticlePageHandling);
        }
    }
}
