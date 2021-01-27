using System;

namespace SharePointPnP.IdentityModel.Extensions.S2S
{
    internal class TokenHelperOptions : ICloneable
    {
        //
        // Hosted app configuration
        //
        public string ClientId { get; set; }
        public string HostedAppHostName { get; set; }
        public string ClientSecret { get; set; }
        public string Realm { get; set; }
        //
        // Environment Constants
        //
        public string AcsHostUrl { get; set; } = "accesscontrol.windows.net";
        public string GlobalEndPointPrefix { get; set; } = "accounts";
        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}
