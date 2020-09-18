using Newtonsoft.Json;
using System;

namespace PnP.Framework.Provisioning.Model.Configuration.Tenant
{
    public class ApplyTenantConfiguration
    {
        [JsonProperty("doNotWaitForSitesToBeFullyCreated")]
        public bool DoNotWaitForSitesToBeFullyCreated { get; set; }

        [JsonIgnore]
        [Obsolete("Use DoNotWaitForSitesToBeFullyCreated")]
        public int DelayAfterModernSiteCreation { get; set; }
    }
}
