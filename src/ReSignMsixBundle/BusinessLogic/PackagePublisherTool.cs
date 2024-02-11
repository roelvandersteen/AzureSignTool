using System.Xml.Linq;
using System.Xml.XPath;

namespace ReSignMsixBundle.BusinessLogic;

[SupportedOSPlatform("windows")] internal static class PackagePublisherTool
{
    /// <summary>Modify the package publisher in the manifests of a list of MSIX files.</summary>
    /// <param name="msixFiles">The MSIX files.</param>
    /// <param name="publisher">The publisher.</param>
    /// <param name="cancellationToken">A token that allows processing to be cancelled.</param>
    public static void ModifyPackagePublisher(IEnumerable<string> msixFiles, string publisher, CancellationToken cancellationToken)
    {
        foreach (var msixFile in msixFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ModifyPackagePublisher(msixFile, publisher);
        }
    }

    private static void ModifyPackagePublisher(string msixFile, string publisher)
    {
        var manifestPath = string.Empty;
        using (var zip = ZipFile.OpenRead(msixFile))
        {
            foreach (var entry in zip.Entries)
            {
                if (!entry.Name.Equals("AppxManifest.xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                manifestPath = Path.Combine(Path.GetTempPath(), entry.Name);
                entry.ExtractToFile(manifestPath, true);

                ModifyPublisherInManifest(manifestPath, publisher);
                break;
            }
        }

        if (string.IsNullOrEmpty(manifestPath))
        {
            throw new InvalidOperationException($"AppxManifest.xml was not found in {msixFile}");
        }

        new AppxPackageEditorWrapper().UpdatePackageManifest(msixFile, manifestPath);
    }

    private static void ModifyPublisherInManifest(string manifestPath, string publisher)
    {
        var xmlDoc = XDocument.Load(manifestPath);
        if (xmlDoc.XPathEvaluate("//*[local-name()='Identity']/@Publisher") is XAttribute publisherAttribute)
        {
            publisherAttribute.Value = publisher;
        }

        xmlDoc.Save(manifestPath);
    }
}
