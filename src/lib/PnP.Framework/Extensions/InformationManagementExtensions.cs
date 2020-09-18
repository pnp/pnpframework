using Microsoft.SharePoint.Client.InformationPolicy;
using PnP.Framework.Entities;
using PnP.Framework.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{

    /// <summary>
    /// Class that deals with information management features
    /// </summary>
    public static partial class InformationManagementExtensions
    {

        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
        public static bool HasSitePolicyApplied(this Web web)
        {
            return Task.Run(() => web.HasSitePolicyAppliedImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
        public static async Task<bool> HasSitePolicyAppliedAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.HasSitePolicyAppliedImplementation();
        }

        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
        private static async Task<bool> HasSitePolicyAppliedImplementation(this Web web)
        {
            var hasSitePolicyApplied = ProjectPolicy.DoesProjectHavePolicy(web.Context, web);
            await web.Context.ExecuteQueryRetryAsync();
            return hasSitePolicyApplied.Value;
        }

        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        public static DateTime GetSiteExpirationDate(this Web web)
        {
            return Task.Run(() => web.GetSiteExpirationDateImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        public static async Task<DateTime> GetSiteExpirationDateAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSiteExpirationDateImplementation();
        }

        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        private static async Task<DateTime> GetSiteExpirationDateImplementation(this Web web)
        {
            if (await web.HasSitePolicyAppliedImplementation())
            {
                var expirationDate = ProjectPolicy.GetProjectExpirationDate(web.Context, web);
                await web.Context.ExecuteQueryRetryAsync();
                return expirationDate.Value;
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
        public static DateTime GetSiteCloseDate(this Web web)
        {
            return Task.Run(() => web.GetSiteCloseDateImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
        public static async Task<DateTime> GetSiteCloseDateAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSiteCloseDateImplementation();
        }

        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
        private static async Task<DateTime> GetSiteCloseDateImplementation(this Web web)
        {
            if (await web.HasSitePolicyAppliedImplementation())
            {
                var closeDate = ProjectPolicy.GetProjectCloseDate(web.Context, web);
                await web.Context.ExecuteQueryRetryAsync();
                return closeDate.Value;
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
        public static List<SitePolicyEntity> GetSitePolicies(this Web web)
        {
            return Task.Run(() => web.GetSitePoliciesImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
        public static async Task<List<SitePolicyEntity>> GetSitePoliciesAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSitePoliciesImplementation();
        }

        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
        private static async Task<List<SitePolicyEntity>> GetSitePoliciesImplementation(this Web web)
        {
            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
            await web.Context.ExecuteQueryRetryAsync();

            var policies = new List<SitePolicyEntity>();

            if (sitePolicies != null && sitePolicies.Count > 0)
            {
                foreach (var policy in sitePolicies)
                {
                    policies.Add(new SitePolicyEntity
                    {
                        Name = policy.Name,
                        Description = policy.Description,
                        EmailBody = policy.EmailBody,
                        EmailBodyWithTeamMailbox = policy.EmailBodyWithTeamMailbox,
                        EmailSubject = policy.EmailSubject
                    });
                }
            }

            return policies;
        }

        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
        public static SitePolicyEntity GetAppliedSitePolicy(this Web web)
        {
            return Task.Run(() => web.GetAppliedSitePolicyImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
        public static async Task<SitePolicyEntity> GetAppliedSitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetAppliedSitePolicyImplementation();
        }

        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
        private static async Task<SitePolicyEntity> GetAppliedSitePolicyImplementation(this Web web)
        {
            if (await web.HasSitePolicyAppliedImplementation())
            {
                var policy = ProjectPolicy.GetCurrentlyAppliedProjectPolicyOnWeb(web.Context, web);
                web.Context.Load(policy,
                             p => p.Name,
                             p => p.Description,
                             p => p.EmailSubject,
                             p => p.EmailBody,
                             p => p.EmailBodyWithTeamMailbox);
                await web.Context.ExecuteQueryRetryAsync();

                return new SitePolicyEntity
                {
                    Name = policy.Name,
                    Description = policy.Description,
                    EmailBody = policy.EmailBody,
                    EmailBodyWithTeamMailbox = policy.EmailBodyWithTeamMailbox,
                    EmailSubject = policy.EmailSubject
                };
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
        public static SitePolicyEntity GetSitePolicyByName(this Web web, string sitePolicy)
        {
            return Task.Run(() => web.GetSitePolicyByNameImplementation(sitePolicy)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
        public static async Task<SitePolicyEntity> GetSitePolicyByNameAsync(this Web web, string sitePolicy)
        {
            await new SynchronizationContextRemover();
            return await web.GetSitePolicyByNameImplementation(sitePolicy);
        }

        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
        private static async Task<SitePolicyEntity> GetSitePolicyByNameImplementation(this Web web, string sitePolicy)
        {
            var policies = await web.GetSitePoliciesAsync();

            if (policies.Count > 0)
            {
                var policy = policies.FirstOrDefault(p => p.Name == sitePolicy);
                return policy;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
        public static bool ApplySitePolicy(this Web web, string sitePolicy)
        {
            return Task.Run(() => web.ApplySitePolicyImplementation(sitePolicy)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
        public static async Task<bool> ApplySitePolicyAsync(this Web web, string sitePolicy)
        {
            await new SynchronizationContextRemover();
            return await web.ApplySitePolicyImplementation(sitePolicy);
        }

        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
        private static async Task<bool> ApplySitePolicyImplementation(this Web web, string sitePolicy)
        {
            var result = false;

            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
            await web.Context.ExecuteQueryRetryAsync();

            if (sitePolicies != null && sitePolicies.Count > 0)
            {
                var policyToApply = sitePolicies.FirstOrDefault(p => p.Name == sitePolicy);

                if (policyToApply != null)
                {
                    ProjectPolicy.ApplyProjectPolicy(web.Context, web, policyToApply);
                    await web.Context.ExecuteQueryRetryAsync();
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
        public static bool IsClosedBySitePolicy(this Web web)
        {
            return Task.Run(() => web.IsClosedBySitePolicyImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
        public static async Task<bool> IsClosedBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.IsClosedBySitePolicyImplementation();
        }

        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
        private static async Task<bool> IsClosedBySitePolicyImplementation(this Web web)
        {
            var isClosed = ProjectPolicy.IsProjectClosed(web.Context, web);
            await web.Context.ExecuteQueryRetryAsync();
            return isClosed.Value;
        }

        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        public static bool SetClosedBySitePolicy(this Web web)
        {
            return Task.Run(() => web.SetClosedBySitePolicyImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        public static async Task<bool> SetClosedBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.SetClosedBySitePolicyImplementation();
        }

        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        private static async Task<bool> SetClosedBySitePolicyImplementation(this Web web)
        {
            if (await web.HasSitePolicyAppliedImplementation() && !await web.IsClosedBySitePolicyImplementation())
            {
                ProjectPolicy.CloseProject(web.Context, web);
                await web.Context.ExecuteQueryRetryAsync();
                return true;
            }
            return false;
        }
        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
        public static bool SetOpenBySitePolicy(this Web web)
        {
            return Task.Run(() => web.SetOpenBySitePolicyImplementation()).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
        public static async Task<bool> SetOpenBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.SetOpenBySitePolicyImplementation();
        }

        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
        private static async Task<bool> SetOpenBySitePolicyImplementation(this Web web)
        {
            if (await web.HasSitePolicyAppliedImplementation() && await web.IsClosedBySitePolicyImplementation())
            {
                ProjectPolicy.OpenProject(web.Context, web);
                await web.Context.ExecuteQueryRetryAsync();
                return true;
            }
            return false;
        }
    }
}
