using Microsoft.SharePoint.Client;
using System;

namespace PnP.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Decimal value into a Double
    /// </summary>
    internal class FromStringToBasePermissionsValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            BasePermissions bp = new BasePermissions();
            var basePermissionString = sourceValue as string;
            // Is it an int value (for backwards compability)?
            int permissionInt;
            if (int.TryParse(basePermissionString, out permissionInt))
            {
                bp.Set((PermissionKind)permissionInt);
            }
            else if (!string.IsNullOrEmpty(basePermissionString))
            {
                foreach (var pk in basePermissionString.Split(','))
                {
                    PermissionKind permissionKind;
                    if (Enum.TryParse<PermissionKind>(pk, out permissionKind))
                    {
                        bp.Set(permissionKind);
                    }
                }
            }
            return bp;
        }
    }
}
