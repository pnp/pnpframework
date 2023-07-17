# Getting started with the PnP Framework

The PnP Framework is a collection of libraries and tools that simplify development for Microsoft 365 and SharePoint Online. 

The PnP Framework is open source and maintained by the SharePoint Patterns and Practices (PnP) community. It is available as a NuGet package and can be used in .NET Framework, .NET Core, and .NET Standard projects.

## What's in the PnP Framework?

Pnp Framework includes a range of capabilities that simplify development for Microsoft 365, SharePoint Online, and SharePoint on-premises. It includes:

- Set of utility methods for working with SharePoint Online
- PnP Modernization Framework
- PnP Provisioning Engine
- Helper classes and extension methods to make it easier to perform standard operations
- Incorporates the PnP Core SDK


## Referencing the PnP Framework in your project

The PnP Framework is available as a NuGet package. You can reference it in your project by adding a reference to the NuGet package.

Package location: [PnP.Framework Package | Nuget](https://www.nuget.org/packages/PnP.Framework)
 
```dotnetcli
dotnet add package PnP.Framework --version 1.13.xx-nightly
```

**Script and polyglot notebooks**

```dotnetcli
#r "nuget: PnP.Framework, 1.13.xx-nightly"
```

## Update Cadence

Each night these preview packages are refreshed, so you can always upgrade to the latest dev bits by upgrading your NuGet package to the latest version.

## If you want to know how PnP Framework is build, here is the code?

The PnP Framework is maintained in the PnP GitHub repository: https://github.com/pnp/pnpframework. You'll find:

- The code of the PnP Framework in the `src\sdk\lib` folder
- The source of the documentation you are reading right now in the `docs` folder

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework" aria-hidden="true" />

