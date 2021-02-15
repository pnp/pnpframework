using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    public class ResponseHeader
    {
        public string SchemaVersion { get; set; } = "15.0.0.0";
        public string LibraryVersion { get; set; } = "16.0.10355.20000";
        public string ErrorInfo { get; set; }
        public string TraceCorrelationId { get; set; } = "3aa6409f-0e7b-10e2-4775-0f0cc16c6f95";
    }
}
