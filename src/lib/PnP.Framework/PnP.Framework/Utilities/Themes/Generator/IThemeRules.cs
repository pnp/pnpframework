using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities.Themes.Generator
{
    public interface IThemeRules : IEnumerable<String>
    {
        IThemeSlotRule this[string key] { get; set; }
    }
}
