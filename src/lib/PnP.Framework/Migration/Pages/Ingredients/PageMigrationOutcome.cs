namespace PnP.Framework.Migration.Pages.Ingredients
{
    public enum PageMigrationOutcome
    {
        Unknown = 0,
        Exact = 1,
        ExecutableWithTransform = 2,
        ExecutableWithLoss = 3,
        /// <summary>
        /// Legacy aggregate outcome. New plans never emit this ambiguous state;
        /// they distinguish mitigation work from literal HTTP authorization stops.
        /// </summary>
        Blocked = 4,

        /// <summary>
        /// One or more ingredients require more evidence, planning, or capability
        /// work. This is a nonterminal work-queue state and must be re-planned.
        /// </summary>
        MitigationPending = 5,

        /// <summary>
        /// One or more ingredient branches have retained literal wire HTTP 401 or
        /// 403 evidence. Only this outcome represents an authorization stop.
        /// </summary>
        AuthorizationBlocked = 6,

        /// <summary>
        /// The proposed action graph is internally inconsistent. This is an RCA
        /// input, not an authorization stop.
        /// </summary>
        Invalid = 7,

        /// <summary>
        /// At least one independent ingredient transaction is executable while
        /// another ingredient branch is deferred or authorization-blocked. The
        /// execution frontier determines the exact transaction boundary.
        /// </summary>
        PartiallyExecutable = 8
    }
}
