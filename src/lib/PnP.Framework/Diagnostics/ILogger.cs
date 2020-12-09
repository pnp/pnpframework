namespace PnP.Framework.Diagnostics
{
    /// <summary>
    /// Interface for Logging
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Log Information
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Info(LogEntry entry);
        /// <summary>
        /// Warning Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Warning(LogEntry entry);
        /// <summary>
        /// Error Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
#pragma warning disable CA1716
        void Error(LogEntry entry);
#pragma warning restore CA1716
        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Debug(LogEntry entry);
    }
}
