using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    internal sealed class ListCalculatedFieldOrderResult
    {
        public IList<ListFieldMaterializationPlan> Fields { get; set; } = new List<ListFieldMaterializationPlan>();

        public IList<string> CycleFields { get; set; } = new List<string>();
    }

    internal static class ListCalculatedFieldOrder
    {
        private static readonly Regex FieldReference = new Regex(@"\[(?<name>[^\]]+)\]", RegexOptions.CultureInvariant);

        public static ListCalculatedFieldOrderResult Order(IEnumerable<ListFieldMaterializationPlan> fields)
        {
            var values = (fields ?? Enumerable.Empty<ListFieldMaterializationPlan>()).ToArray();
            var calculated = values.Where(IsCalculated).ToArray();
            var byName = new Dictionary<string, ListFieldMaterializationPlan>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in calculated)
            {
                AddName(byName, field.InternalName, field);
                AddName(byName, field.Title, field);
            }
            var dependencies = calculated.ToDictionary(value => value, value => new HashSet<ListFieldMaterializationPlan>());
            var dependents = calculated.ToDictionary(value => value, value => new List<ListFieldMaterializationPlan>());
            foreach (var field in calculated)
            {
                foreach (var reference in ReadFormulaReferences(field.SourceSchemaXml))
                {
                    ListFieldMaterializationPlan dependency;
                    if (byName.TryGetValue(reference, out dependency) && !ReferenceEquals(dependency, field)
                        && dependencies[field].Add(dependency))
                    {
                        dependents[dependency].Add(field);
                    }
                }
            }
            var ready = new SortedSet<ListFieldMaterializationPlan>(
                dependencies.Where(value => value.Value.Count == 0).Select(value => value.Key),
                Comparer<ListFieldMaterializationPlan>.Create((left, right) =>
                {
                    var compared = StringComparer.OrdinalIgnoreCase.Compare(left.InternalName, right.InternalName);
                    return compared != 0 ? compared : left.SourceFieldId.CompareTo(right.SourceFieldId);
                }));
            var orderedCalculated = new List<ListFieldMaterializationPlan>();
            while (ready.Count > 0)
            {
                var current = ready.Min;
                ready.Remove(current);
                orderedCalculated.Add(current);
                foreach (var dependent in dependents[current])
                {
                    dependencies[dependent].Remove(current);
                    if (dependencies[dependent].Count == 0)
                    {
                        ready.Add(dependent);
                    }
                }
            }
            var cycle = calculated.Except(orderedCalculated).OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).ToArray();
            return new ListCalculatedFieldOrderResult
            {
                Fields = values.Where(value => !IsCalculated(value))
                    .OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                    .Concat(orderedCalculated)
                    .Concat(cycle)
                    .ToList(),
                CycleFields = cycle.Select(value => value.InternalName).ToList()
            };
        }

        private static bool IsCalculated(ListFieldMaterializationPlan value)
        {
            return value.Disposition == ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated;
        }

        private static void AddName(IDictionary<string, ListFieldMaterializationPlan> values, string name, ListFieldMaterializationPlan field)
        {
            if (!string.IsNullOrWhiteSpace(name) && !values.ContainsKey(name))
            {
                values[name] = field;
            }
        }

        private static IEnumerable<string> ReadFormulaReferences(string schemaXml)
        {
            if (string.IsNullOrWhiteSpace(schemaXml))
            {
                return Enumerable.Empty<string>();
            }
            try
            {
                var formula = XDocument.Parse(schemaXml).Descendants()
                    .FirstOrDefault(value => string.Equals(value.Name.LocalName, "Formula", StringComparison.OrdinalIgnoreCase))?.Value;
                return string.IsNullOrWhiteSpace(formula)
                    ? Enumerable.Empty<string>()
                    : FieldReference.Matches(formula).Cast<Match>().Select(value => value.Groups["name"].Value.Trim()).Distinct(StringComparer.OrdinalIgnoreCase);
            }
            catch (System.Xml.XmlException)
            {
                return Enumerable.Empty<string>();
            }
        }
    }
}
