using PnP.Framework.Utilities.Themes.Palettes;
using System.Collections.Generic;

namespace PnP.Framework.Utilities.Themes.Generator
{
    public interface IThemeSlotRule
    {
        /* The name of this theme slot. */
        string name { get; set; }
        /* The actual color this theme slot is if it is a color. */
        IColor color { get; set; }
        /* The value of this slot if it is NOT a color. Must be falsey if not a color. */
        string value { get; set; }
        /* The theme slot this slot is based on. */
        IThemeSlotRule inherits { get; set; }
        /* If set, this slot is the specified shade of the slot it inherits from. */
        Shade asShade { get; set; }
        /* Whether this slot is a background shade, which uses different logic for generating its inheriting-as-shade value. */
        bool isBackgroundShade { get; set; }
        /* Whether this slot has been manually overridden (else, it was automatically generated based on inheritance). */
        bool isCustomized { get; set; }
        /* A collection of rules that inherit from this one. It is the responsibility of the inheriting rule to add itself to its parent's dependentRules collection. */
        List<IThemeSlotRule> dependentRules { get; set; }
    }
}
