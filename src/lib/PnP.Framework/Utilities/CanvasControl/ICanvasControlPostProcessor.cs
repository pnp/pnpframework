using Microsoft.SharePoint.Client;

namespace PnP.Framework.Utilities.CanvasControl
{
    /// <summary>
    ///     Interface for WebPart Post Processing
    /// </summary>
    public interface ICanvasControlPostProcessor
    {
        /// <summary>
        ///     Method for processing canvas control
        /// </summary>
        /// <param name="canvasControl">Canvas control object</param>
        /// <param name="context">ClientContext to use</param>
        void Process(Framework.Provisioning.Model.CanvasControl canvasControl, ClientContext context);
    }
}