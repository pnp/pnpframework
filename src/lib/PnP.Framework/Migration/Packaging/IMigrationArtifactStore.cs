using System.IO;

namespace PnP.Framework.Migration.Packaging
{
    public interface IMigrationArtifactStore
    {
        bool Contains(string sha256);

        Stream OpenRead(string sha256);

        ArtifactReference Put(Stream content, string mediaType = null, string originalName = null);
    }
}
