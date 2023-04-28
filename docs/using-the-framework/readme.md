# Getting started with the PnP Framework

> [!Note]
> This is draft content , we are reviewing and preparing up this article

## Where is the code?

The PnP Framework is maintained in the PnP GitHub repository: https://github.com/pnp/pnpframework. You'll find:

- The code of the PnP Framework in the `src\sdk\lib` folder
- The source of the documentation you are reading right now in the `docs` folder

## Referencing the PnP Framework in your project

The PnP Framework is available as a NuGet package. You can reference it in your project by adding a reference to the NuGet package.

 - [Nuget Package](https://www.nuget.org/packages/PnP.Framework)
 
```dotnetcli
dotnet add package PnP.Framework --version 1.12.29-nightly
```

Script and polyglot notebooks

```dotnetcli
#r "nuget: PnP.Framework, 1.12.29-nightly"
```

## Update Cadence

Each night these preview packages are refreshed, so you can always upgrade to the latest dev bits by upgrading your NuGet package to the latest version.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework" aria-hidden="true" />