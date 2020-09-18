using System;

namespace PnP.Framework.Attributes
{
    [AttributeUsage(AttributeTargets.Class,
                       AllowMultiple = true)]
    public sealed class TokenDefinitionDescriptionAttribute : Attribute
    {
        public string Token { get; set; }
        public string Description { get; set; }
        public string Returns { get; set; }
        public string Example { get; set; }
    }
}
