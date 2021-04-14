using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.PnPSdk
{
    internal class PnPCoreSdkAuthenticationProviderFactory : ILegacyAuthenticationProviderFactory
    {
        public ILegacyAuthenticationProvider GetAuthenticationProvider(ClientContext context)
        {
            return new PnPCoreSdkAuthenticationProvider(context);
        }
    }
}
