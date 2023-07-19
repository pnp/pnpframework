# Working with Microsoft SharePoint with PnP Framework

The following samples show how to work with Microsoft SharePoint using the PnP Framework.

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

- [Getting Started with the PnP Framework | PolyGlot Notebook](https://github.com/pnp/pnpframework/tree/dev/docs/notebooks/Getting-Started-with-PnP-Framework.ipynb)


> [!Note]
> These are a work in progress to provider richer samples, but these take a while to write. If you have any suggestions or wish to contribute your examples, please raise an issue in the GitHub repository.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/working-with-sharepoint" aria-hidden="true" />