using System;

namespace PnP.Framework.Graph.Model
{
    /// <summary>
    /// Represents a single value of a Directory Setting
    /// </summary>
    public class DirectorySettingValue
    {
        public string DefaultValue { get; set; }

        public string Description { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public string Value { get; set; }
    }
}
