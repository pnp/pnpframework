using PnP.Framework.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class ContentTypeBindingsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.ContentTypeBinding>();
            var contentTypeBindings = source.GetPublicInstancePropertyValue("ContentTypeBindings");

            if (contentTypeBindings != null)
            {
                var bindingResolvers = new Dictionary<string, IResolver>
                {
                    { $"{typeof(Model.FieldRef).FullName}.Id", new FromStringToGuidValueResolver() }
                };

                foreach (var binding in (IEnumerable)contentTypeBindings)
                {
                    var targetItem = new Model.ContentTypeBinding();
                    PnPObjectsMapper.MapProperties(binding, targetItem, bindingResolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return result;
        }
    }
}
