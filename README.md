# PnP Framework

**PnP Framework** is a .NET Standard 2.0 / .NET 5.0 library targeting Microsoft 365 containing the PnP Provisioning engine and a ton of other useful extensions. This library is the cross platform successor of the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library. The original [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library only works on Windows (due to the dependency on .Net Framework) and supports SharePoint on-premises and SharePoint Online, while this library will work cross platform but only supports SharePoint Online. Going forward we'll only be **actively maintaining PnP Framework** and once PnP Framework is declared GA we'll retire the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library.

Next to PnP Framework that will be replacing [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) we're also building a brand new [PnP Core SDK](https://github.com/pnp/pnpcore) that targets modern .Net development and which will be working everywhere where .NET will run. This library is the long term evolution of PnP Framework, but we'll enable a phased transition from PnP Framework to [PnP Core SDK](https://github.com/pnp/pnpcore).

## Getting started

Pull down the latest version of PnP Framework here: [![PnP.Framework Nuget package](https://img.shields.io/nuget/vpre/PnP.Framework.svg)](https://www.nuget.org/packages/PnP.Framework/).

## PnP .Net roadmap status

We've shipped our first PnP Framework preview version and preview 3 of the [PnP Core SDK](https://github.com/pnp/pnpcore).

![PnP dotnet roadmap](PnP%20dotnet%20Roadmap%20-%20December%20status.png)

## I've found a bug, where do I need to log an issue or create a PR

Between now and the end of 2020 both [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) and PnP Framework are actively maintained. Once PnP Framework GA's we'll stop maintaining [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core).

Given PnP Framework is our future going forward we would prefer issues and PR's being created in the PnP Framework repo. If you want your PR to apply to both then it's recommended to create the PR in both repositories for the time being.

## Building and contributing

To build PnP Framework you need the following minimal components installed:

- [Visual Studio 2019 version 16.8+](https://visualstudio.microsoft.com/vs/)
- [.NET SDK version 5.0](https://dotnet.microsoft.com/download/dotnet/5.0)

Contributions should be made against the **dev** branch of the repository.

**Community rocks, sharing is caring!**

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
