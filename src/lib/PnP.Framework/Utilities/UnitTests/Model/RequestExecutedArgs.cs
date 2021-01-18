using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    public class RequestExecutedArgs : EventArgs
    {
        public string RequestBody { get; set; }
        public string ResponseBody { get; set; }
        public string CalledUrl { get; set; }
    }
}
