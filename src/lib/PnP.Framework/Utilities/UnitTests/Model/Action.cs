using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    public class BaseAction
    {
        [XmlAttribute]
        public int Id { get; set; }
        [XmlAttribute]
        public string ObjectPathId { get; set; }
        public virtual string GetResponse<T>(Identity identity, List<MockResponseEntry<T>> responseEntries, CSOMRequest request)
        {
            return "{\"IsNull\":false}";
        }
    }
    [XmlInclude(typeof(BaseAction))]
    public class IdentityQueryAction : BaseAction
    {
        public override string GetResponse<T>(Identity identity, List<MockResponseEntry<T>> responseEntries, CSOMRequest request)
        {
            if (identity is ObjectPathMethod)
            {
                ObjectPathMethod objectPathMethod = identity as ObjectPathMethod;
                MockResponseEntry<T> contextEntry = responseEntries.FirstOrDefault(entry =>
                       entry.Method == objectPathMethod.Name);
                if (contextEntry == null)
                {
                    return base.GetResponse<T>(identity, responseEntries, request);
                }
                return contextEntry.SerializedReturnValue;
            }
            return $"{{\"_ObjectIdentity_\":\"{identity.Name}\"}}";
        }
    }
    [XmlInclude(typeof(BaseAction))]
    public class QueryAction : BaseAction
    {
        [XmlElement(ElementName = "Query")]
        public SelectQuery SelectQuery { get; set; }

        public override string GetResponse<T>(Identity identity, List<MockResponseEntry<T>> responseEntries, CSOMRequest request)
        {
            if (identity is ObjectPathMethod)
            {
                ObjectPathMethod objectPathMethod = identity as ObjectPathMethod;
                Identity parentIdentity = request.ObjectPaths.FirstOrDefault(parent => parent.Id == objectPathMethod.ParentId);
                List<string> parentMethodParameter = new List<string>();
                if (parentIdentity is ObjectPathMethod)
                {
                    ObjectPathMethod parentMethod = parentIdentity as ObjectPathMethod;
                    if (parentMethod.Parameters.Count > 0)
                    {
                        parentMethodParameter.AddRange(parentMethod.Parameters.Select(param => param.Value as String));
                    }
                }
                //TODO: refactor it as after one day I don't understand it!
                MockResponseEntry<T> contextEntry = responseEntries.First(entry =>
                    entry.Method == objectPathMethod.Name
                    && (parentMethodParameter.Count == 0 ||
                        entry.ParentParameterValues.All(parentValue => parentMethodParameter.Any(parentParam => parentParam == parentValue))
                    && entry.NameValueParameters.All(
                        methodParam => (objectPathMethod.Parameters.Any(
                            param => param.Value == methodParam.Value)) ||
                            objectPathMethod.Parameters.Any(param => param.Properties.Any(paramProp => paramProp.Value == methodParam.Value && paramProp.Name == methodParam.Key)))));

                return contextEntry.SerializedReturnValue;
            }
            else if (identity is Property)
            {
                Property associatedProperty = identity as Property;

                MockResponseEntry<T> contextEntry = responseEntries.FirstOrDefault(entry =>
                    entry.PropertyName == associatedProperty.Name);
                if(contextEntry == null)
                {
                    Identity parentIdenity = request.ObjectPaths.First(obj => obj.Id == associatedProperty.ParentId);

                    return GetResponse<T>(parentIdenity, responseEntries, request);
                }
                return contextEntry.SerializedReturnValue;
            }
            else if (identity is StaticMethod)
            {
                StaticMethod associatedProperty = identity as StaticMethod;
                MockResponseEntry<T> contextEntry = responseEntries.First(entry =>
                    entry.Method == associatedProperty.Name);
                return contextEntry.SerializedReturnValue;
            }
            else if (identity is ObjectPathMethod)
            {
                ObjectPathMethod pathMethod = identity as ObjectPathMethod;
                MockResponseEntry<T> contextEntry = responseEntries.First(entry =>
                    entry.Method == pathMethod.Name);
                return contextEntry.SerializedReturnValue;
            }
            else if (identity is Identity)
            {
                MockResponseEntry<T> contextEntry = responseEntries.FirstOrDefault(entry =>
                    entry.PropertyName == identity.Name);
                if (contextEntry != null)
                {
                    return contextEntry.SerializedReturnValue;
                }
            }
            return base.GetResponse<T>(identity, responseEntries, request);
        }
    }
    public class Parameter
    {
        [XmlAttribute]
        public string Type { get; set; }
        [XmlAttribute]
        public string Name { get; set; }
        [XmlText]
        public string Value { get; set; }
    }
    public class SelectQuery
    {
        [XmlAttribute]
        public bool SelectAllProperties { get; set; }
        public List<Property> Properties { get; set; }
    }
    [XmlInclude(typeof(BaseAction))]
    public class MethodAction : BaseAction
    {
        [XmlAttribute]
        public string Name { get; set; }
        public List<Parameter> Parameters { get; set; }
        public override string GetResponse<T>(Identity identity, List<MockResponseEntry<T>> responseEntries, CSOMRequest request)
        {
            MockResponseEntry<T> contextEntry = responseEntries.First(entry =>
                entry.Method == this.Name
                && entry.NameValueParameters.All(methodParam => this.Parameters.Any(param => param.Value == methodParam.Value)));
            if (contextEntry.ReturnValue == null)
            {

            }
            return contextEntry.SerializedReturnValue;
        }
    }
}
