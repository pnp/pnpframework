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

## Using PolyGlot Notebooks

The PolyGlot samples provide working examples of how to interact with SharePoint, that can be ran against your Microsoft 365 tenant, written in C#.


```dotnetcli
#r "nuget: PnP.Framework, 1.13.xx-nightly"
```


The simplest way to get working samples, is to check out the notebooks, please navigate to the GitHub repository, the following notebooks are available, to use these, you will need to *clone the repository* and open the notebooks in Visual Studio Code, using the [Extension - PolyGlot Notebooks by Microsoft](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.dotnet-interactive-vscode).


Examples:

- [Getting Started with the PnP Framework | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Framework.ipynb)
- [Getting Started with the PnP Provisioning | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Provisioning.ipynb)


> [!Note]
> These are a work in progress to provider richer samples, but these take a while to write. If you have any suggestions or wish to contribute your examples, please raise an issue in the GitHub repository.

If you feel samples should be present in the documentation, please [raise an issue in the GitHub repository](https://github.com/pnp/pnpframework/issues), and lets discuss your thoughts.


## If you want to know how PnP Framework is build, here is the code?

The PnP Framework is maintained in the PnP GitHub repository: https://github.com/pnp/pnpframework. You'll find:

- The code of the PnP Framework in the `src\sdk\lib` folder
- The source of the documentation you are reading right now in the `docs` folder

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework" aria-hidden="true" />

