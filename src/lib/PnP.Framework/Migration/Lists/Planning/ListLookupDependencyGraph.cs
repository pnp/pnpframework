using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    public sealed class ListLookupDependency
    {
        public Guid SourceListId { get; set; }

        public Guid LookupListId { get; set; }

        public Guid FieldId { get; set; }

        public string FieldInternalName { get; set; }
    }

    public sealed class ListDependencyOrderResult
    {
        public IList<Guid> OrderedSourceListIds { get; set; } = new List<Guid>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsExecutable => Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error);
    }

    public static class ListLookupDependencyGraph
    {
        public static ListDependencyOrderResult Order(IEnumerable<Guid> sourceListIds, IEnumerable<ListLookupDependency> lookupDependencies)
        {
            if (sourceListIds == null)
            {
                throw new ArgumentNullException(nameof(sourceListIds));
            }
            if (lookupDependencies == null)
            {
                throw new ArgumentNullException(nameof(lookupDependencies));
            }

            var nodes = sourceListIds.Where(value => value != Guid.Empty).Distinct().OrderBy(value => value).ToArray();
            var nodeSet = new HashSet<Guid>(nodes);
            var edges = lookupDependencies
                .Where(value => nodeSet.Contains(value.SourceListId) && nodeSet.Contains(value.LookupListId))
                .GroupBy(value => value.SourceListId.ToString("D") + ":" + value.LookupListId.ToString("D"), StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            var indegree = nodes.ToDictionary(value => value, value => 0);
            var dependents = nodes.ToDictionary(value => value, value => new List<Guid>());
            foreach (var edge in edges)
            {
                indegree[edge.SourceListId]++;
                dependents[edge.LookupListId].Add(edge.SourceListId);
            }

            var ready = new SortedSet<Guid>(indegree.Where(value => value.Value == 0).Select(value => value.Key));
            var ordered = new List<Guid>();
            while (ready.Count > 0)
            {
                var current = ready.Min;
                ready.Remove(current);
                ordered.Add(current);
                foreach (var dependent in dependents[current].OrderBy(value => value))
                {
                    indegree[dependent]--;
                    if (indegree[dependent] == 0)
                    {
                        ready.Add(dependent);
                    }
                }
            }

            var result = new ListDependencyOrderResult { OrderedSourceListIds = ordered };
            if (ordered.Count != nodes.Length)
            {
                var cycle = indegree.Where(value => value.Value > 0).Select(value => value.Key).OrderBy(value => value).ToArray();
                result.Issues.Add(new MigrationIssue
                {
                    Code = "LookupDependencyCycle",
                    Severity = MigrationIssueSeverity.Blocker,
                    Subject = "list-lookup-graph",
                    Ingredient = "List.LookupClosure",
                    Message = "Lookup dependency cycle detected among source Lists: " + string.Join(", ", cycle.Select(value => value.ToString("D"))) + "."
                });
            }
            return result;
        }
    }
}
