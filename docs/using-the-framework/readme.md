# Getting started with the PnP Framework

The PnP Framework is a collection of libraries and tools that simplify development for Microsoft 365 and SharePoint Online. 

The PnP Framework is open source and maintained by the SharePoint Patterns and Practices (PnP) community. It is available as a NuGet package and can be used in .NET Framework, .NET Core, and .NET Standard projects.

## What's in the PnP Framework?

Pnp Framework includes a range of capabilities that simplify development for Microsoft 365, SharePoint Online, and SharePoint on-premises. It includes:

- Set of utility methods to extend CSOM 
- PnP Modernization Framework
- PnP Provisioning Engine
- Helper classes and extension methods to make it easier to perform standard operations
- Incorporates the PnP Core SDK

## Images in modernized rich text

Modernization converts images in wiki content, publishing rich-text fields, and
non-script Content Editor Web Part (CEWP) HTML (embedded or linked) to CK5 inline
images. Images stay in text controls, including tables and lists, without image-driven
splitting or duplicate Image web parts. Links, alt text, and supported sizing are
preserved; HTML may be normalized. CK4 targets are unsupported.

After the target page is saved, external HTTPS images are imported through
SharePoint's `AddImageFromExternalUrl` service into its page-owned asset folder.
Repeated transformations reuse existing assets for the same complete external URL
(including query parameters), even across processes, without re-downloading or
refreshing them. Tenant restrictions apply; recognized stock-image URLs remain external.

External import or conversion failures retain the URL in a **native CK5 image** with a warning.
These fallbacks depend on external access and browser policies, and lack SharePoint
file-backed features. Unsupported URLs and internal-image conversion failures retain
raw image HTML with a warning; CK5 editing may remove it.

`AddTableListImageAsImageWebPart` is ignored. `HandleWikiImagesAndVideos` controls only
iframe/video placement. Standalone Image web parts and script/embed CEWPs are unchanged.

Verify image placement and links after editing and saving the migrated page.

## I don't have access to a Microsoft 365 tenant

If you don't have a Microsoft 365 tenant you can, for developer purposes, always request [a free developer tenant](https://developer.microsoft.com/en-us/microsoft-365/dev-program) and use that for developing and testing your applications. 

> Note: When your organization already uses Microsoft 365 it's still a good practice to develop and test your applications on a non-production tenant.

## Referencing the PnP Framework in your project

The PnP Framework is available as a NuGet package. You can reference it in your project by adding a reference to the NuGet package.

Package location: [PnP.Framework Package | Nuget](https://www.nuget.org/packages/PnP.Framework)
 
```dotnetcli
dotnet add package PnP.Framework --version 1.13.xx-nightly
```

## Update Cadence

Each night these preview packages are refreshed, so you can always upgrade to the latest dev bits by upgrading your NuGet package to the latest version.

## Samples use PolyGlot Notebooks

The PolyGlot samples provide working examples of how to interact with SharePoint, that can be ran against your Microsoft 365 tenant, written in C# and runnable code - as an aid to understand and learning of how to use the PnP Framework.

The simplest way to get working samples, is to check out the notebooks in our repo, please navigate to the [GitHub repository](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks), the following notebooks are available, to use these, you will need to *clone the repository* and open the notebooks in Visual Studio Code, using the [Extension - PolyGlot Notebooks by Microsoft](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.dotnet-interactive-vscode).

To get setup, please follow the instructions in the [Setting up to use Polyglot samples](setting-up-to-use-polyglot-samples.md) guide.

Examples:

- [Getting Started with the PnP Framework | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Framework.ipynb)
- [Getting Started with the PnP Provisioning | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Provisioning.ipynb)


_*Note: These are a work in progress to provider richer samples, but these take a while to write. If you have any suggestions or wish to contribute your examples, please raise an issue in the GitHub repository._

If you feel samples should be present in the documentation, please [raise an issue in the GitHub repository](https://github.com/pnp/pnpframework/issues), and lets discuss your thoughts.


## If you want to know how PnP Framework is built, here is the code?

The PnP Framework is maintained in the PnP GitHub repository: https://github.com/pnp/pnpframework. You'll find:

- The code of the PnP Framework in the `src\sdk\lib` folder
- The source of the documentation you are reading right now in the `docs` folder

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework" aria-hidden="true" />
