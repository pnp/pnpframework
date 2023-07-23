# Setting up to run the Polyglot Notebooks

The Polyglot Notebooks are a great way to get started with the PnP Framework, with runnable samples and you can adjust the code and explore the framework. There are a few things you need to do to get started.

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)
- [Install .Net 7 SDK](https://dotnet.microsoft.com/en-us/download)
- [.NET Interactive Notebooks Extension](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.dotnet-interactive-vscode)

## Configuring the Notebooks for your tenant

The Polyglot Notebooks are designed to be run against your tenant, but you will need to configure these settings file to do this.

### Settings File

To keep credentials away from the PolyGlot Notebooks, we have setup a seperate file to contain these, copy the `appsettings.sample.json` file to `appsettings.json` and update the values to match your environment.

```json
{
    "azureAppId":"<app-id>",
    "certificatePassword":"<cert-password>",
    "certificatePath":"C:\\temp\\PolyGlot\\pnpframework-polyglot.pfx",
    "azureTenantName":"contoso.onmicrosoft.com",
    "siteUrl" : "https://contoso.sharepoint.com/sites/contoso"
}

```

You will need to create an Azure AD App Registration and generate a certificate, assign permissions to SharePoint within the app. For these notebooks, we have kept access to this for the time being.


*There is more coming on how to do this.*

## Running the Notebooks

...
