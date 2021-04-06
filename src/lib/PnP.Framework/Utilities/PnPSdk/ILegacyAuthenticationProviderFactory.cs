using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Text;
using PnP.Core.Services;

namespace PnP.Framework.Utilities.PnPSdk
{
    internal interface ILegacyAuthenticationProviderFactory
    {
        ILegacyAuthenticationProvider GetAuthenticationProvider(ClientContext context);
    }
}
