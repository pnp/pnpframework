# Build Notes

Run the following:
C:\Git\docfx\2.59.4\docfx metadata docfx.json

Working version:
```
C:\Git\docfx\2.59.3\docfx metadata docfx.json
C:\Git\docfx\2.59.3\docfx build docfx.json
```

# Errors to fix

[22-12-03 06:14:29.540]Warning:[BuildCore.Build Document.CompilePhaseHandlerWithIncremental.TocDocumentProcessor.Prebuild.BuildTocDocument](toc.yml)Unable to find either toc.yml or toc.md inside using-the-sdk/. Make sure the file is included in config file docfx.json!
[22-12-03 06:14:29.540]Warning:[BuildCore.Build Document.CompilePhaseHandlerWithIncremental.TocDocumentProcessor.Prebuild.BuildTocDocument](toc.yml)Unable to find either toc.yml or toc.md inside tutorials/. Make sure the file is included in config file docfx.json!
[22-12-03 06:14:29.555]Warning:[BuildCore.Build Document.CompilePhaseHandlerWithIncremental.TocDocumentProcessor.Prebuild.BuildTocDocument](toc.yml)Unable to find either toc.yml or toc.md inside contributing/. Make sure the file is included in config file docfx.json!


# Documentation Plan

## Using the framework

The following pages are being considered for creation:

- Examples of creating the client context
- Show examples of working with SharePoint List Items and Documents
- Describe the extension methods areas
- Examples of working with the Microsoft Graph
- Examples of working with Microsoft Teams
- Examples of working with Modernization Tooling
- Examples of working with the PnP Provisioning Engine
- Examples of working with PnP Core SDK within the PnP Framework

Examples are also console or Polyglot working examples to provide working code samples.