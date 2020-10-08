# PnP Framework

**PnP Framework** is a .Net Standard 2.0 library targeting Microsoft 365 containing the PnP Provisioning engine and a ton of other useful extensions. This library is the cross platform successor of the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library. The original [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library only works on Windows (due to the dependency on .Net Framework) and supports SharePoint on-premises and SharePoint Online, while this library will work cross platform but only supports SharePoint Online. Going forward we'll only be **actively maintaining PnP Framework** and once PnP Framework is declared GA we'll retire the [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) library.

Next to PnP Framework that will be replacing [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) we're also building a brand new [PnP Core SDK](https://github.com/pnp/pnpcore) that targets modern .Net development (including .Net 5) and which will be working everywhere where .Net will run. This library is the long term evolution of PnP Framework, but we'll enable a phased transition from PnP Framework to [PnP Core SDK](https://github.com/pnp/pnpcore).

## PnP .Net roadmap status

We've shipped our first PnP Framework preview version and preview 3 of the [PnP Core SDK](https://github.com/pnp/pnpcore).

![PnP dotnet roadmap](PnP%20dotnet%20Roadmap%20-%20October%20status.png)

## I've found a bug, where do I need to log an issue or create a PR

Between now and the end of 2020 both [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) and PnP Framework are actively maintained. Once PnP Framework GA's we'll stop maintaining [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core).

Given PnP Framework is our future going forward we would prefer issues and PR's being created in the PnP Framework repo. If you want your PR to apply to both then it's recommended to create the PR in both repositories for the time being.
