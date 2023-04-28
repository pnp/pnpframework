# Writing Tests

## Before you get started

Running unit tests is a great way to validate that your changes are working as expected.  The unit tests are located in the `./src/lib/PnP.Framework.Test` folder.
Ensure that you have gone through the steps in the [Setup](setup.md) article.

## Writing tests

There are a series of existing tests that you can use as a reference; all the tests rely on an element of setup in the tenant and you will need a live tenant to run this against.
In the `./src/lib/PnP.Framework.Test` folder you will find a series of folders that contain the tests for each of the different areas of the framework.  For example, the `./src/lib/PnP.Framework.Test/Authentication` folder contains the tests for the authentication providers.

> [!Warning]
> Do not run the tests against a production tenant, some of the tests access tenant wide features that could impact on existing setup.