using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes.Execution
{
    internal static class ContentTypeClosureMaterializer
    {
        public static void Ensure(
            ClientContext anchorContext,
            IEnumerable<ContentTypeClosureNodePlan> plans,
            MigrationExecutionRecorder recorder)
        {
            var values = (plans ?? Enumerable.Empty<ContentTypeClosureNodePlan>())
                .GroupBy(value => value.TargetOwnerWebUrl + "\u001f" + value.Schema.ContentTypeId, StringComparer.OrdinalIgnoreCase)
                .Select(group =>
                {
                    var candidates = group.ToArray();
                    if (candidates.Select(value => value.PlanDigest).Distinct(StringComparer.OrdinalIgnoreCase).Count() != 1)
                    {
                        throw new InvalidOperationException("Conflicting content type closure plans target '" + group.Key + "'.");
                    }
                    return candidates[0];
                })
                .OrderBy(value => value.Schema.ContentTypeId.Length)
                .ThenBy(value => value.Schema.ContentTypeId, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (values.Length == 0)
            {
                recorder.RecordAlreadySatisfied("schema.content-types", "The approved List closure has no custom site content types.");
                return;
            }

            foreach (var plan in values)
            {
                using (var context = anchorContext.Clone(plan.TargetOwnerWebUrl))
                {
                    var probe = ContentTypeTargetInspector.Inspect(context, context.Web, plan.Schema);
                    var admission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan.Schema, probe);
                    if (!admission.IsEligible)
                    {
                        throw new InvalidOperationException("Fresh target content type preflight failed for '" + plan.Schema.ContentTypeId + "': "
                            + string.Join("; ", admission.Issues.Select(value => value.Message)));
                    }
                    recorder.Execute(
                        "schema.content-type." + plan.SourceOwnerWebId.ToString("N") + "." + plan.Schema.ContentTypeId,
                        "Ensure site content type '" + plan.Schema.Name + "' (" + plan.Schema.ContentTypeId + ").",
                        () => ContentTypeMaterializer.Ensure(context, context.Web, plan.Schema, admission),
                        disposition => disposition == ContentTypeMaterializationDisposition.ReuseOwned
                            ? MutationOutcome.AlreadySatisfied
                            : MutationOutcome.Applied,
                        disposition => "Site content type disposition: " + disposition + ".");
                }
            }
        }
    }
}
