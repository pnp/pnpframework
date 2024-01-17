# Getting Started

The PnP Provisioning Engine is a set of libraries and tools that enable the creation and deployment of site templates and artifacts in SharePoint Online and SharePoint on-premises. The engine includes a set of .NET libraries that can be used to programmatically create and apply templates, as well as a console application that can be used to apply templates from the command line. The templates can include site columns, content types, lists, libraries, pages, and other SharePoint artifacts. 

Key features of PnP Provisioning include:

- Create templates from existing SharePoint sites
- Apply templates to existing SharePoint sites
- Site level or tenant level provisioning
- Pass parameters to the templates
- Captures a wide range of SharePoint artifacts


## PnP Provisioning Schema

The PnP Provisioning templates are based on an XML schema that captures the structure and content of a SharePoint site. The schema is extensible and can be used to capture a wide range of SharePoint artifacts. The schema is located here: [Schema 09/2022 | PnP Provisioning Schema](https://github.com/pnp/PnP-Provisioning-Schema/blob/master/ProvisioningSchema-2022-09.md)


## Using the Polyglot Notebook samples

The PolyGlot samples provide working examples of how to interact with SharePoint, that can be ran against your Microsoft 365 tenant, written in C#.

### Settings File

To keep credentials away from the PolyGlot Notebooks, we have setup a seperate file to contain these, copy the `appsettings.sample.json` file to `appsettings.json` and update the values to match your environment.

```json
{
    "azureAppId":"<app-id>",
    "certificatePassword":"<cert-password>",
    "certificatePath":"C:\\temp\\PolyGlot\\pnpframework-polyglot.pfx",
    "azureTenantName":"contoso.onmicrosoft.com"
}

```

You will need to create an Azure AD App Registration and generate a certificate, assign permissions to SharePoint within the app. For these notebooks, we have kept access to this for the time being.

*There is more coming on how to do this.*

## Samples available

The simplest way to get to the notebooks, is to navigate to the GitHub repository, the following notebooks are available, to use these, you will need to *clone the repository* and open the notebooks in Visual Studio Code, using the [Extension - PolyGlot Notebooks by Microsoft](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.dotnet-interactive-vscode).

Examples:

- [Getting Started with the PnP Provisioning | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Provisioning.ipynb)


> [!Note]
> These are a work in progress to provider richer samples, but these take a while to write. If you have any suggestions or wish to contribute your examples, please raise an issue in the GitHub repository.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/pnp-provisioning" aria-hidden="true" />