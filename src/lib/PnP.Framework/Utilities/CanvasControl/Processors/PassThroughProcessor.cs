using Microsoft.SharePoint.Client;

namespace PnP.Framework.Utilities.CanvasControl.Processors
{
    /// <summary>
    ///     Default processor when others are not resolved
    /// </summary>
    public class CanvasControlPassThroughProcessor : ICanvasControlPostProcessor
    {
        /// <summary>
        ///  Method for processing canvas control
        /// </summary>
        /// <param name="canvasControl">Canvas control object</param>
        /// <param name="context">ClientContext to use</param>
        public void Process(Framework.Provisioning.Model.CanvasControl canvasControl, ClientContext context)
        {
        }
    }
}