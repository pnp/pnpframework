using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a collection type from Domain Model to Schema
    /// </summary>
    internal class RoleAssigmentsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public RoleAssigmentsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            List<RoleAssignment> res = new List<RoleAssignment>();
            var sourceValue = source.GetPublicInstancePropertyValue("RoleAssignment");
            if (sourceValue != null)
            {
                res = PnPObjectsMapper.MapObjects(sourceValue, new CollectionFromSchemaToModelTypeResolver(typeof(RoleAssignment)), null, true) as List<RoleAssignment>;
            }
            return res;
        }
    }
}
