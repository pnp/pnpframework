using System.Resources;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("PnP.PowerShell")]

#if DEBUG
[assembly: InternalsVisibleTo("PnP.Framework.Test")]
[assembly: InternalsVisibleTo("PnP.Framework.Modernization.Test")]
#endif

[assembly: NeutralResourcesLanguage("en")]