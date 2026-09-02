using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public static class PageIngredientPlanEvaluator
    {
        public static PageIngredientPlanEvaluation Evaluate(
            CanonicalPageIngredientGraph graph,
            IEnumerable<PageIngredientAction> actions)
        {
            var issues = new List<MigrationIssue>();
            if (graph == null)
            {
                issues.Add(Issue("IngredientGraphMissing", "ingredient-graph", "The canonical page ingredient graph is missing."));
                return Result(PageMigrationOutcome.Blocked, issues);
            }

            var nodes = graph.Nodes ?? new List<PageIngredientNode>();
            var actionList = (actions ?? Array.Empty<PageIngredientAction>()).ToList();
            var nodeById = UniqueNodes(nodes, issues);
            var actionByIngredient = UniqueActions(actionList, issues);

            foreach (var node in nodes.Where(value => value != null && value.HasContent))
            {
                if (!actionByIngredient.TryGetValue(node.Id, out var action))
                {
                    issues.Add(Issue("IngredientDispositionMissing", node.Id, "Every nonempty ingredient must have exactly one planned disposition."));
                    continue;
                }

                if (action.Disposition == IngredientDisposition.Undefined)
                {
                    issues.Add(Issue("IngredientDispositionUndefined", node.Id, "The ingredient action has no semantic disposition."));
                }

                if (action.Disposition == IngredientDisposition.Block)
                {
                    issues.Add(Issue("IngredientBlocked", node.Id, action.Reason ?? "The ingredient is blocked."));
                }

                if (action.Capability == IngredientCapability.Unknown
                    && action.Disposition != IngredientDisposition.Drop
                    && action.Disposition != IngredientDisposition.Delegate
                    && action.Disposition != IngredientDisposition.Block)
                {
                    issues.Add(Issue("IngredientCapabilityUnknown", node.Id, "A retained ingredient has unknown target capability."));
                }
            }

            foreach (var action in actionList.Where(value => value != null))
            {
                if (string.IsNullOrWhiteSpace(action.IngredientId) || !nodeById.ContainsKey(action.IngredientId))
                {
                    issues.Add(Issue("IngredientActionOrphaned", action.IngredientId ?? "ingredient-action", "The action does not reference a captured ingredient."));
                }
            }

            ValidateDependencyReleases(graph.Edges, nodeById, actionByIngredient, issues);
            ValidateRequiredEdges(graph.Edges, nodeById, actionByIngredient, issues);
            if (issues.Any(value => value.Severity == MigrationIssueSeverity.Blocker || value.Severity == MigrationIssueSeverity.Error))
            {
                return Result(PageMigrationOutcome.Blocked, issues);
            }

            var materialActions = actionList.Where(value => value != null
                && nodeById.TryGetValue(value.IngredientId ?? string.Empty, out var node)
                && node.HasContent).ToList();
            if (materialActions.Any(value => value.Disposition == IngredientDisposition.Drop || value.Disposition == IngredientDisposition.Delegate))
            {
                return Result(PageMigrationOutcome.ExecutableWithLoss, issues);
            }

            if (materialActions.Any(value => value.Disposition == IngredientDisposition.Transform || value.Disposition == IngredientDisposition.Substitute))
            {
                return Result(PageMigrationOutcome.ExecutableWithTransform, issues);
            }

            return Result(PageMigrationOutcome.Exact, issues);
        }

        private static Dictionary<string, PageIngredientNode> UniqueNodes(
            IEnumerable<PageIngredientNode> nodes,
            ICollection<MigrationIssue> issues)
        {
            var result = new Dictionary<string, PageIngredientNode>(StringComparer.Ordinal);
            foreach (var node in nodes)
            {
                if (node == null || string.IsNullOrWhiteSpace(node.Id) || result.ContainsKey(node.Id))
                {
                    issues.Add(Issue("IngredientIdentityInvalid", node?.Id ?? "ingredient", "Ingredient IDs must be nonempty and unique."));
                    continue;
                }

                result.Add(node.Id, node);
            }
            return result;
        }

        private static Dictionary<string, PageIngredientAction> UniqueActions(
            IEnumerable<PageIngredientAction> actions,
            ICollection<MigrationIssue> issues)
        {
            var result = new Dictionary<string, PageIngredientAction>(StringComparer.Ordinal);
            foreach (var action in actions)
            {
                if (action == null || string.IsNullOrWhiteSpace(action.IngredientId) || result.ContainsKey(action.IngredientId))
                {
                    issues.Add(Issue("IngredientActionIdentityInvalid", action?.IngredientId ?? "ingredient-action", "Each ingredient may have exactly one action."));
                    continue;
                }

                result.Add(action.IngredientId, action);
            }
            return result;
        }

        private static void ValidateRequiredEdges(
            IEnumerable<PageIngredientEdge> edges,
            IDictionary<string, PageIngredientNode> nodes,
            IDictionary<string, PageIngredientAction> actions,
            ICollection<MigrationIssue> issues)
        {
            foreach (var edge in edges ?? Array.Empty<PageIngredientEdge>())
            {
                if (edge == null
                    || !nodes.ContainsKey(edge.FromIngredientId ?? string.Empty)
                    || !nodes.ContainsKey(edge.ToIngredientId ?? string.Empty))
                {
                    issues.Add(Issue("IngredientEdgeInvalid", edge?.FromIngredientId ?? "ingredient-edge", "Every graph edge must connect two captured ingredients."));
                    continue;
                }

                actions.TryGetValue(edge.FromIngredientId, out var consumer);
                var explicitlyReleased = consumer?.Disposition == IngredientDisposition.Transform
                    && consumer.ReleasedDependencyIngredientIds != null
                    && consumer.ReleasedDependencyIngredientIds.Contains(edge.ToIngredientId, StringComparer.Ordinal);
                if (edge.Requirement != PageIngredientRequirement.Required
                    || consumer == null
                    || !IsRetained(consumer.Disposition)
                    || !nodes[edge.ToIngredientId].HasContent
                    || explicitlyReleased)
                {
                    continue;
                }

                if (!actions.TryGetValue(edge.ToIngredientId, out var dependency) || !IsRetained(dependency.Disposition))
                {
                    issues.Add(Issue(
                        "RequiredIngredientDependencyUnsatisfied",
                        edge.FromIngredientId,
                        $"Retained ingredient '{edge.FromIngredientId}' requires '{edge.ToIngredientId}', but the dependency is not retained or explicitly released by a transform."));
                }
            }
        }

        private static void ValidateDependencyReleases(
            IEnumerable<PageIngredientEdge> edges,
            IDictionary<string, PageIngredientNode> nodes,
            IDictionary<string, PageIngredientAction> actions,
            ICollection<MigrationIssue> issues)
        {
            var requiredEdges = new HashSet<string>(
                (edges ?? Array.Empty<PageIngredientEdge>())
                    .Where(value => value != null && value.Requirement == PageIngredientRequirement.Required)
                    .Select(value => EdgeIdentity(value.FromIngredientId, value.ToIngredientId)),
                StringComparer.Ordinal);
            foreach (var action in actions.Values.Where(value => value != null))
            {
                var releases = action.ReleasedDependencyIngredientIds ?? Array.Empty<string>();
                var duplicate = releases
                    .GroupBy(value => value, StringComparer.Ordinal)
                    .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
                if (duplicate != null)
                {
                    issues.Add(Issue(
                        "IngredientDependencyReleaseInvalid",
                        action.IngredientId,
                        "Released dependency IDs must be nonempty and unique."));
                }
                if (releases.Count > 0 && action.Disposition != IngredientDisposition.Transform)
                {
                    issues.Add(Issue(
                        "IngredientDependencyReleaseInvalid",
                        action.IngredientId,
                        "Only a Transform action may explicitly release a required dependency."));
                }
                foreach (var dependencyId in releases.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal))
                {
                    if (!nodes.ContainsKey(dependencyId)
                        || !requiredEdges.Contains(EdgeIdentity(action.IngredientId, dependencyId)))
                    {
                        issues.Add(Issue(
                            "IngredientDependencyReleaseInvalid",
                            action.IngredientId,
                            $"Released ingredient '{dependencyId}' is not a captured required dependency of '{action.IngredientId}'."));
                    }
                }
            }
        }

        private static bool IsRetained(IngredientDisposition disposition)
        {
            return disposition == IngredientDisposition.Preserve
                || disposition == IngredientDisposition.Transform
                || disposition == IngredientDisposition.Substitute;
        }

        private static string EdgeIdentity(string from, string to)
        {
            return (from ?? string.Empty) + "\u001f" + (to ?? string.Empty);
        }

        private static MigrationIssue Issue(string code, string ingredient, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = ingredient,
                Ingredient = ingredient,
                Message = message
            };
        }

        private static PageIngredientPlanEvaluation Result(PageMigrationOutcome outcome, IList<MigrationIssue> issues)
        {
            return new PageIngredientPlanEvaluation
            {
                Outcome = outcome,
                Issues = issues
            };
        }
    }
}
