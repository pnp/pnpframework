# PnP Framework documentation

**PnP Framework** is a .NET Standard 2.0 / .NET 5.0 / .NET 6.0 / .NET 7.0 library targeting Microsoft 365 containing the PnP Provisioning engine and a ton of other useful extensions. This library is the cross platform successor of the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library. 

## History and Roadmap

The original [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library only works on Windows (due to the dependency on .Net Framework) and supports SharePoint on-premises and SharePoint Online, while this library will work cross platform but only supports SharePoint Online. Going forward we'll only be **actively maintaining PnP Framework**, the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library is retired and archived.

Next to PnP Framework that will be replacing [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) we're also building a brand new [PnP Core SDK](https://github.com/pnp/pnpcore) that targets modern .NET development and which will work everywhere where .NET will run. This library is the long term evolution of PnP Framework, we'll enable a phased transition from PnP Framework to [PnP Core SDK](https://github.com/pnp/pnpcore) without impacting PnP Framework users.

## Getting started

Pull down the latest version of PnP Framework here:

Nuget package | Description | Latest release | Latest nightly development version
--------------|-------------|----------------|------------------------------------
PnP.Framework | The PnP Framework library | [![PnP.Framework Nuget package](https://img.shields.io/nuget/v/PnP.Framework.svg)](https://www.nuget.org/packages/PnP.Framework/) | [![PnP.Framework Nuget package](https://img.shields.io/nuget/vpre/PnP.Framework.svg)](https://www.nuget.org/packages/PnP.Framework/)

### YouTube Content

To help you understand the PnP.Framework, we have the following YouTube content to help you learn:

- [Using PnP Libraries to simplify SharePoint solution design and implementation](https://www.youtube.com/watch?v=MLHMjJNHehE)
- [SharePoint site provisioning with Microsoft Teams Notifications using Bot Framework
](https://www.youtube.com/watch?v=x7nrzafsbAw)

## Migrating from PnP Sites Core

This library is the cross platform successor of the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core), see the [Migrating from PnP Sites Core to PnP Framework guide](using-the-framework/migrating-from-pnp-sites-core.md) to learn how to switch over.

## I've found a bug, where do I need to log an issue or create a PR

Issues and PR's need to be created in the PnP Framework repo, the PnP Sites Core repo has been archived.

## Supportability and SLA

This library is open-source and community provided component with active community providing support for it. This is not Microsoft provided component so there's no SLA or direct support for this open-source component from Microsoft. Please report any issues using the [issues list](https://github.com/pnp/pnpframework/issues).

## Building and contributing

There is guidance in this [contributing](contributing/readme.md) part of the document, where you can learn to:

- Setup the Repo and the Tooling
- Writing documentation
- writing tests
- Getting recognised

**Community rocks, sharing is caring!**

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.