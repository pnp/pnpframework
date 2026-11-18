namespace PnP.Framework
{
    /// <summary>The kind of identity represented by an authentication provider.</summary>
    public enum AuthenticationIdentityType
    {
        /// <summary>The identity type is not known.</summary>
        Unknown,
        /// <summary>A delegated user identity.</summary>
        User,
        /// <summary>An application identity.</summary>
        Application
    }
}
