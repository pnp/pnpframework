using System;
using System.Collections.Generic;

namespace PnP.Framework.Utilities.Themes.Generator
{
    public interface IThemeRules : IEnumerable<String>
    {
        IThemeSlotRule this[string key] { get; set; }
    }
}
