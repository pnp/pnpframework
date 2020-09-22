using System;

namespace PnP.Framework.Modernization.Functions
{
    /// <summary>
    /// Exception class thrown when a SharePoint object (e.g. list) is not available at the target site
    /// </summary>
    public class NotAvailableAtTargetException: Exception
    {
        /// <summary>
        /// Throws a NotAvailableAtTargetException message
        /// </summary>
        /// <param name="message">Error message</param>
        public NotAvailableAtTargetException(string message): base(message) { }

        /// <summary>
        /// Throws a NotAvailableAtTargetException message
        /// </summary>
        /// <param name="message">Error message</param>
        /// <param name="innerException">Inner exception object</param>
        public NotAvailableAtTargetException(string message, Exception innerException) : base(message, innerException) { }
    }
}
