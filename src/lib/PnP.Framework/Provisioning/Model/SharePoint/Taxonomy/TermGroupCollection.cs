using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of TermGroup objects
    /// </summary>
    public partial class TermGroupCollection : BaseProvisioningTemplateObjectCollection<TermGroup>
    {
        /// <summary>
        /// Constructor for TermGroupCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TermGroupCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
