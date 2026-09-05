using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal sealed class MarkdownReportWriter
    {
        private readonly StringBuilder builder = new StringBuilder();

        public void Heading(int level, string value)
        {
            builder.AppendLine(new string('#', level) + " " + value);
            builder.AppendLine();
        }

        public void Paragraph(string value)
        {
            builder.AppendLine(value);
            builder.AppendLine();
        }

        public void Table(string heading, string[] headers, IEnumerable<string[]> rows)
        {
            if (!string.IsNullOrWhiteSpace(heading))
            {
                Heading(2, heading);
            }

            builder.AppendLine("| " + string.Join(" | ", headers.Select(EscapeTableCell)) + " |");
            builder.AppendLine("| " + string.Join(" | ", headers.Select(_ => "---")) + " |");
            var any = false;
            foreach (var row in rows ?? Array.Empty<string[]>())
            {
                any = true;
                builder.AppendLine("| " + string.Join(" | ", row.Select(EscapeTableCell)) + " |");
            }

            if (!any)
            {
                builder.AppendLine("| " + string.Join(" | ", headers.Select((_, index) => index == 0 ? "None" : string.Empty)) + " |");
            }

            builder.AppendLine();
        }

        public void List(string heading, IEnumerable<string> values)
        {
            Heading(2, heading);
            var items = (values ?? Array.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            if (items.Length == 0)
            {
                builder.AppendLine("- None");
            }
            else
            {
                foreach (var item in items)
                {
                    builder.AppendLine($"- {item}");
                }
            }

            builder.AppendLine();
        }

        public override string ToString()
        {
            return builder.ToString();
        }

        private static string EscapeTableCell(string value)
        {
            return PublishingPageReportValueFormatter.Format(value)
                .Replace("|", "\\|")
                .Replace("\r", " ")
                .Replace("\n", " ");
        }
    }
}
