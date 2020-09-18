using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities.CanvasControl.Processors;

namespace PnP.Framework.Utilities.CanvasControl
{
    public class CanvasControlPostProcessorFactory
    {
        /// <summary>
        /// Resolves client control web part by type
        /// </summary>
        /// <param name="canvasControl">CanvasControl object</param>
        /// <returns>Returns PassThroughProcessor object</returns>
        public static ICanvasControlPostProcessor Resolve(Framework.Provisioning.Model.CanvasControl canvasControl)
        {
            if (canvasControl.Type == WebPartType.List)
            {
                return new ListControlPostProcessor(canvasControl);
            }

            return new CanvasControlPassThroughProcessor();
        }
    }
}