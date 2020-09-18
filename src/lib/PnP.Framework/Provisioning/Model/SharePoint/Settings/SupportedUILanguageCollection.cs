using PnP.Framework.Extensions;
using System;
using System.Linq;

namespace PnP.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SupportedUILanguage objects
    /// </summary>
    public partial class SupportedUILanguageCollection : BaseProvisioningTemplateObjectCollection<SupportedUILanguage>, IEquatable<SupportedUILanguageCollection>
    {
        /// <summary>
        /// Constructor for SupportedUILangaugeCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SupportedUILanguageCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }

        /// <summary>
        /// Compare languages
        /// </summary>
        /// <param name="other">Collection of languages to compare with</param>
        /// <returns>True if the same, false otherwise</returns>
        public bool Equals(SupportedUILanguageCollection other)
        {
            if (other == null)
            {
                return (false);
            }

            return this.Items.AsEnumerable<SupportedUILanguage>().DeepEquals(other.Items.AsEnumerable<SupportedUILanguage>());
        }
    }
}
