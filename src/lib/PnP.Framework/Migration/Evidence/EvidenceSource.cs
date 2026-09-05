namespace PnP.Framework.Migration.Evidence
{
    public sealed class EvidenceSource
    {
        public string ExchangeId { get; set; }

        public string PayloadSha256 { get; set; }

        public string Selector { get; set; }
    }
}
