using System;
using System.Xml.Linq;

namespace PnP.Framework.Tests.Framework.Functional.Validators
{
    public class ValidateXmlEventArgs : EventArgs
    {
        public XElement SourceObject { get; set; }
        public XElement TargetObject { get; set; }
        public bool IsEqual { get; set; }


        public ValidateXmlEventArgs(XElement sourceObject, XElement targetObject)
        {
            SourceObject = sourceObject;
            TargetObject = targetObject;
            IsEqual = false;
        }

    }
}
