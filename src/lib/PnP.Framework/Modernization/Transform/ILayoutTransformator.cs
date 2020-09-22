using PnP.Framework.Modernization.Entities;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Modernization.Transform
{
    /// <summary>
    /// Interface implemented by all layout transformators
    /// </summary>
    public interface ILayoutTransformator
    {
        /// <summary>
        /// Transforms a classic wiki/webpart page layout into a modern client side page layout
        /// </summary>
        /// <param name="pageData">Information about the analyed page</param>
        void Transform(Tuple<Pages.PageLayout, List<WebPartEntity>> pageData);
    }
}
