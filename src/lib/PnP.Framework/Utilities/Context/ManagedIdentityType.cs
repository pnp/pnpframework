namespace PnP.Framework.Utilities.Context
{
    /// <summary>
    /// Types of Managed Identity supported within the Framework
    /// </summary>
    public enum ManagedIdentityType
    {
        /// <summary>
        /// System Assigned Managed Identity
        /// </summary>
        SystemAssigned = 0,

        /// <summary>
        /// User Assigned Managed Identity, referenced by its client Id
        /// </summary>
        UserAssignedByClientId = 1,

        /// <summary>
        /// User Assigned Managed Identity, referenced by its object Id
        /// </summary>
        UserAssignedByObjectId = 2,

        /// <summary>
        /// User Assigned Managed Identity, refernced by its resource Id
        /// </summary>
        UserAssignedByResourceId = 3
    }
}
