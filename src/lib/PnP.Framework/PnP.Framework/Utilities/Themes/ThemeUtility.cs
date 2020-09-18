using PnP.Framework.Utilities.Themes.Generator;
using PnP.Framework.Utilities.Themes.Palettes;
using System;

namespace PnP.Framework.Utilities.Themes
{
    public static class ThemeUtility
    {
        /// <summary>
        /// Apply three custom colors to the standard Office UI Fabric template and get back the JSON with the CSS rules
        /// </summary>
        /// <param name="primaryColor">Primary Color for the Theme</param>
        /// <param name="bodyTextColor">Body Text Color for the Theme</param>
        /// <param name="bodyBackgroundColor">Body Background Color for the Theme</param>
        /// <returns>THe JSON representation of the Theme</returns>
        public static String GetThemeAsJSON(String primaryColor, String bodyTextColor, String bodyBackgroundColor)
        {
            // Validate inputs
            if (String.IsNullOrEmpty(primaryColor))
                throw new ArgumentNullException(nameof(primaryColor));

            if (String.IsNullOrEmpty(bodyTextColor))
                throw new ArgumentNullException(nameof(bodyTextColor));

            if (String.IsNullOrEmpty(bodyBackgroundColor))
                throw new ArgumentNullException(nameof(bodyBackgroundColor));

            // Parse the custom colors
            var primaryCustomColor = Colors.getColorFromString(primaryColor);
            var bodyTextCustomColor = Colors.getColorFromString(bodyTextColor);
            var bodyBackgroundCustomColor = Colors.getColorFromString(bodyBackgroundColor);

            // Get the standard template
            var standardRules = new ThemeRulesStandard();
            ThemeGenerator.InsureSlots(standardRules, false);

            // Set the custom colors to the template
            ThemeGenerator.SetSlot(standardRules[BaseSlots.primaryColor.ToString()], primaryCustomColor);
            ThemeGenerator.SetSlot(standardRules[BaseSlots.foregroundColor.ToString()], bodyTextCustomColor);
            ThemeGenerator.SetSlot(standardRules[BaseSlots.backgroundColor.ToString()], bodyBackgroundCustomColor);

            // Get the JSON string
            String json = ThemeGenerator.GetThemeAsJson(standardRules);

            return json;
        }
    }

}
