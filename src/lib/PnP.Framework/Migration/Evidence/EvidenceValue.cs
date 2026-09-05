using System.Collections.Generic;

namespace PnP.Framework.Migration.Evidence
{
    public sealed class EvidenceValue<T>
    {
        public EvidenceAvailability Availability { get; set; }

        public EvidencePresence Presence { get; set; }

        public T Value { get; set; }

        public IList<EvidenceSource> Sources { get; set; } = new List<EvidenceSource>();

        public IList<string> Diagnostics { get; set; } = new List<string>();

        public bool HasValue => Availability == EvidenceAvailability.Captured
            && (Presence == EvidencePresence.Present || Presence == EvidencePresence.Empty);
    }
}
