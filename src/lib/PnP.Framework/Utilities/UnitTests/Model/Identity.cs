using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    public class Identity
    {
        [XmlAttribute]
        public string Id { get; set; }
        [XmlAttribute]
        public string Name { get; set; }

        public virtual MockResponseEntry CreateMockResponse(object associatedResponse, BaseAction action, string calledUrl)
        {
            if (action is MethodAction)
            {
                MethodAction methodAction = action as MethodAction;
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                foreach (Parameter parameter in methodAction.Parameters)
                {
                    if (!String.IsNullOrEmpty(parameter.Value))
                    {
                        parameters.Add(parameter.Value, parameter.Value);
                    }
                }
                return new MockResponseEntry()
                {
                    Url = calledUrl,
                    Method = methodAction.Name,
                    ReturnValue = associatedResponse,
                    NameValueParameters = parameters,
                };
            }
            return null;
        }
    }
    [XmlInclude(typeof(Identity))]
    public class Property : Identity
    {
        [XmlAttribute]
        public string ParentId { get; set; }
        public override MockResponseEntry CreateMockResponse(object associatedResponse, BaseAction action, string calledUrl)
        {
            return new MockResponseEntry()
            {
                Url = calledUrl,
                PropertyName = this.Name,
                ReturnValue = associatedResponse
            };
        }
    }
    [XmlInclude(typeof(Property))]
    public class StaticProperty : Identity
    {
        [XmlAttribute]
        public string TypeId { get; set; }
        public override MockResponseEntry CreateMockResponse(object associatedResponse, BaseAction action, string calledUrl)
        {
            return new MockResponseEntry()
            {
                Url = calledUrl,
                PropertyName = this.Name,
                ReturnValue = associatedResponse
            };
        }
    }
    [XmlInclude(typeof(Property))]
    public class ObjectPathMethod : Property
    {
        [XmlArrayItem(ElementName = "Parameter")]
        public List<MethodParameter> Parameters { get; set; }
        public override MockResponseEntry CreateMockResponse(object associatedResponse, BaseAction action, string calledUrl)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            foreach (MethodParameter parameter in Parameters)
            {
                if (!String.IsNullOrEmpty(parameter.Value))
                {
                    parameters.Add(parameter.Value, parameter.Value);
                }
                else
                {
                    foreach (Parameter property in parameter.Properties)
                    {
                        parameters.Add(property.Name, property.Value);
                    }
                }
            }
            return new MockResponseEntry()
            {
                Url = calledUrl,
                Method = this.Name,
                ReturnValue = associatedResponse,
                NameValueParameters = parameters,
            };
        }
    }
    public class MethodParameter
    {
        [XmlAttribute]
        public string TypeId { get; set; }
        [XmlElement(ElementName = "Property", Type = typeof(Parameter))]
        public List<Parameter> Properties { get; set; }
        [XmlText]
        public string Value { get; set; }
    }
    [XmlInclude(typeof(Identity))]
    public class StaticMethod : Identity
    {
        [XmlAttribute]
        public string TypeId { get; set; }
    }
}
