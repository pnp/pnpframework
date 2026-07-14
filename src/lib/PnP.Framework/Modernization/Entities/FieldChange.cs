using System;

namespace PnP.Framework.Modernization.Entities
{
    /// <summary>
    /// Represents a change detected in a SharePoint list item field value
    /// </summary>
    [Serializable]
    public class FieldChange
    {
        /// <summary>
        /// Gets or sets the internal name of the field that has changed
        /// </summary>
        public string FieldInternalName { get; set; }

        /// <summary>
        /// Gets or sets the new value for the field
        /// </summary>
        public object NewValue { get; set; }

        /// <summary>
        /// Gets or sets the current value of the field (optional, for context)
        /// </summary>
        public object CurrentValue { get; set; }

        /// <summary>
        /// Initializes a new instance of the FieldChange class
        /// </summary>
        public FieldChange()
        {
        }

        /// <summary>
        /// Initializes a new instance of the FieldChange class
        /// </summary>
        /// <param name="fieldInternalName">Internal name of the field</param>
        /// <param name="newValue">New value for the field</param>
        public FieldChange(string fieldInternalName, object newValue)
        {
            FieldInternalName = fieldInternalName;
            NewValue = newValue;
        }

        /// <summary>
        /// Initializes a new instance of the FieldChange class
        /// </summary>
        /// <param name="fieldInternalName">Internal name of the field</param>
        /// <param name="newValue">New value for the field</param>
        /// <param name="currentValue">Current value of the field</param>
        public FieldChange(string fieldInternalName, object newValue, object currentValue)
        {
            FieldInternalName = fieldInternalName;
            NewValue = newValue;
            CurrentValue = currentValue;
        }
    }
}