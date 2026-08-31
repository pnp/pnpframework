namespace PnP.Framework.Migration.PublishingPages
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
