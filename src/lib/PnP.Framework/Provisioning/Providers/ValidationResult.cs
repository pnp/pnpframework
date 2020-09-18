using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Providers
{
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<Exception> Exceptions { get; set; }
    }
}
