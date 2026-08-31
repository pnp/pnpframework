using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PublishingPageMigrationReport
    {
        public string Summary { get; set; }

        public IList<string> CapturedIngredients { get; set; } = new List<string>();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
