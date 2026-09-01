namespace PnP.Framework.Migration.PublishingPages.References
{
    public enum PageReferenceDisposition
    {
        PreserveExternal = 0,
        RewriteToTarget = 1,
        MaterializeAtTarget = 2,
        Delegate = 3,
        Block = 4
    }
}
