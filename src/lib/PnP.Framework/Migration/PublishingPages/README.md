# Publishing-page migration

`PnP.Framework.Migration.PublishingPages` provides a staged framework for migrating classic SharePoint publishing pages while retaining enough source evidence to review the current decision and recover additional data in a later version.

The first profile is Enterprise Wiki. Enterprise Wiki is an entry profile, not the owner of every participating object: fields, Web Parts, dependencies, lifecycle, security, packages, reports, and verification have reusable namespaces of their own.

## Workflow

```text
source connection
    -> capture and seal source snapshot
    -> inspect target and generate plan/report
    -> human reviews and approves plan digest
    -> import exactly that sealed plan
    -> fresh target readback and receipt
```

The stages have deliberately different connection requirements:

| Stage | Source connection | Target connection | Writes target |
| --- | ---: | ---: | ---: |
| Discovery | Yes | No | No |
| Export | Yes | No | No |
| Plan | No | Yes | No |
| Review | No | No | No |
| Import | No | Yes | Yes |
| Verification | No | Yes | Reads after import |

Export therefore remains pipeable and portable: the target is selected only when a migration plan is created.

## Namespace layout

The physical folders and namespaces follow the same ownership boundaries.

| Namespace suffix | Responsibility |
| --- | --- |
| `Capture` | Source capture aggregate, capture policy, capture status, and source stability fence. |
| `Content` | Approved text replacements and publishing-content transformation. |
| `Fields` | Complete list-item field evidence, value serialization, planning dispositions, and field writes. |
| `WebParts` | Shared Web Part snapshots, portability policy, and source export. |
| `Lifecycle` | Source lifecycle evidence and the conservative target lifecycle policy. |
| `References` | Authored URL/dependency discovery, target actions, and payload materialization. |
| `Security` | Permission-inheritance and role-assignment evidence. |
| `Planning` | Target-specific plan, planning options, and migration operation. |
| `Packaging` | Versioned JSON envelopes, canonical serialization, digests, and validation. |
| `Reporting` | Complete human-readable Markdown review output. |
| `Verification` | Target probe model, storage assertions, and fresh-readback verification. |
| `EnterpriseWiki` | Enterprise Wiki discovery, profile policy, export, target inspection, planning, import, and file storage facade. |

Only page-wide identity and path rules remain in the root namespace. New object-specific types should go into the owning domain rather than accumulating in the root folder.

## Public workflow entry points

The Enterprise Wiki profile currently exposes these orchestration APIs:

| Type | Responsibility |
| --- | --- |
| `EnterpriseWikiPageDiscovery` | Find and classify source Enterprise Wiki pages. |
| `EnterpriseWikiPackageExporter` | Capture and seal a source-only export package. |
| `EnterpriseWikiMigrationPlanner` | Inspect a target and produce a sealed migration package plus review report. |
| `EnterpriseWikiMigrationImporter` | Validate approval, execute the plan, and return a fresh-readback receipt. |
| `EnterpriseWikiPackageFileStore` | Save and load export packages, migration packages, receipts, and Markdown reports. |

Generic contract and infrastructure types remain under the object-domain namespaces described above. A future publishing-page profile should compose those domains rather than copy the Enterprise Wiki implementation.

## Versioned artifacts

JSON uses camel-case property names, string enum values, explicit nulls, and case-sensitive property names. The current schemas are:

| Artifact | Schema |
| --- | --- |
| Source export | `pnp-publishing-page-export/v1` |
| Target-specific migration package | `pnp-publishing-page-migration-package/v1` |
| Import receipt | `pnp-publishing-page-import-receipt/v1` |

A breaking JSON contract change requires a new schema version. Moving a CLR type between namespaces does not by itself change the JSON field names, but released CLR API compatibility must still be considered separately.

### Source export envelope

`PublishingPageExportPackage` contains:

| JSON field | Meaning |
| --- | --- |
| `schemaVersion` | Export contract version. |
| `exportedAtUtc` | Time at which source capture completed. |
| `snapshot` | Complete source evidence described below. |
| `snapshotDigest` | SHA-256 over canonical serialization of the complete snapshot. |

`PublishingPageCaptureBundle` contains:

| JSON field | Meaning |
| --- | --- |
| `sourceProfile` | Profile that classified the source, currently `EnterpriseWiki`. |
| `capturePolicy` | Normalized source page path, Web Part inclusion choice, and dependency size limit. |
| `source` | Web, file, list item, content type, version, title, and layout identity. |
| `publishingPageContent` | Full source `PublishingPageContent` HTML. |
| `publishingPageContentSha256` | Digest of the captured HTML. |
| `fields` | Every returned Pages-library field definition and typed/best-effort raw value. |
| `webParts` | Shared Web Part export XML and placement when enabled. |
| `dependencies` | Authored references and captured payloads when safely obtainable. |
| `security` | Inheritance flag and role-assignment evidence. |
| `lifecycle` | Checkout, file level, moderation, created, and modified evidence. |
| `sourceFence` | File ID, version, size, and modified time checked before and after capture. |
| `blockers` | Source conditions that make a later plan non-executable. |
| `warnings` | Source conditions that require review but do not by themselves block planning. |

The source fence detects a page that changed while it was being captured. It is not a distributed lock and does not prevent later source edits.

### Migration package envelope

`PublishingPageMigrationPackage` embeds the complete snapshot and adds:

| JSON field | Meaning |
| --- | --- |
| `schemaVersion` | Migration package contract version. |
| `plannedAtUtc` | Time at which target analysis and sealing completed. |
| `exportSchemaVersion` / `exportedAtUtc` | Provenance of the embedded source export. |
| `state` | `ApprovalReady` when no blocker exists, otherwise `Blocked`. |
| `plan` | Target-specific decisions and assertions. |
| `snapshotDigest` | Must still match the embedded snapshot. |
| `planDigest` | SHA-256 over all target decisions; this is the approval token. |
| `report` | Compact report metadata used to generate the complete Markdown view. |

`PublishingPageMigrationPlan` records:

| JSON field | Meaning |
| --- | --- |
| `sourceSnapshotDigest` | Binds this plan to one exact source snapshot. |
| `sourceWebUrl` / `sourcePageServerRelativeUrl` | Source boundary used by reviewed mappings. |
| `targetWebUrl` / `targetWebServerRelativeUrl` / `targetPageServerRelativeUrl` | Exact approved target. |
| `pageLayoutName` | Target publishing layout selected by the profile. |
| `operation` | Currently `CreatePage`. |
| `targetLifecycle` / `lifecycleReason` | Derived Draft/Published result and explanation. |
| `createOnly` | Currently required to be `true`. |
| `planningPolicy` | Normalized policy inputs copied into the sealed plan. |
| `targetProbe` | Target template, Pages library, layout, lifecycle, page, and dependency observations. |
| `fieldActions` | Exactly one decision for every captured field. |
| `dependencyActions` | Exactly one decision for every captured dependency. |
| `replacements` | Explicit source-to-target text substitutions included in the digest. |
| `expectedPublishingPageContentSha256` | Expected content digest after approved replacements. |
| `storageAssertions` / `browserAssertions` | Required post-import evidence. |
| `blockers` / `warnings` | Review findings. `isExecutable` is derived from an empty blocker list. |

Import requires an exact `approvedPlanDigest` match. Editing any sealed target decision without regenerating the digest makes the package invalid.

### Import receipt

`PublishingPageImportReceipt` records start/completion times, approved digest, target identity, target content type and version, expected and actual lifecycle, expected and persisted content digests, Web Part and dependency counts, per-field write results, warnings, fresh-readback status, and overall success.

The receipt is evidence of what the current importer observed after writing. It is not a rollback journal.

## Complete field capture, selective restore

Field capture and field restore intentionally have different scopes.

Every returned source Pages-library field is captured with:

- field ID, internal name, title, type, schema XML, and read-only/hidden/required flags;
- a typed representation for supported values;
- best-effort runtime type, text, JSON, or Base64 evidence;
- capture status and diagnostics.

The plan then creates exactly one `PageFieldAction` for every captured field. This preserves future recoverability while keeping the current importer conservative.

| Disposition | Interpretation |
| --- | --- |
| `Apply` | Recognized, non-empty, writable, target-present, type-compatible, and supported; Import writes it. |
| `AlreadyHandled` | Page creation/content/layout logic owns the field. |
| `SkipEmpty` | No source value needs restoring. |
| `SkipReadOnly` | SharePoint owns the source or target field. |
| `SkipCalculated` | SharePoint recomputes the value. |
| `TargetFieldMissing` | Recognized source field is absent at the target. |
| `TargetTypeMismatch` | Source and target field types differ. |
| `RequiresMapping` | User, lookup, or taxonomy identity cannot be copied safely across sites. |
| `EvidenceOnly` | Snapshot retains the complete evidence, but no current importer owns it. |
| `CaptureUnavailable` | Definition was captured, but no restorable value was returned. |
| `Block` | The plan cannot execute. |

The Enterprise Wiki profile recognizes a small reviewed set of publishing metadata fields. Unknown fields are not discarded and are not guessed into the target.

## Lifecycle policy

There is no top-level `publish` Boolean input. Lifecycle is derived from captured evidence:

- return `Published` only when source file level is `Published`, checkout type is `None`, and moderation is absent or approved (`0`);
- return `Draft` for every other or contradictory state.

Examples from captured pages used by the focused tests:

| Source evidence | Target result | Reason |
| --- | --- | --- |
| `level=Draft`, `checkOutType=Online`, `moderationStatus=3` | `Draft` | The source is not an unambiguous published version. |
| `level=Published`, `checkOutType=None`, `moderationStatus=0` | `Published` | All available evidence agrees that the source is published. |

If any planned field write fails, the importer avoids publishing the page and records a warning.

## Dependencies and Web Parts

Authored references are inventoried separately from page HTML so each one receives an explicit target action:

| Disposition | Behavior |
| --- | --- |
| `PreserveExternal` | Leave an allowed external reference unchanged. |
| `RewriteToTarget` | Rewrite a same-tenant or same-web reference to the reviewed target location. |
| `MaterializeAtTarget` | Upload a captured same-web payload before page creation. |
| `Delegate` | Reserved for another reviewed migration owner. |
| `Block` | Stop import until the unsupported reference is resolved. |

Same-tenant iframes, resources outside the captured web boundary, and missing payloads block the exact profile. Capturable dependencies are limited by `maximumDependencyBytes`, currently 10 MiB per dependency by default.

Shared Web Parts retain export XML, ID, title, zone, index, hidden state, and digest. Empty/invalid exports, unavailable Web Parts, RSS Aggregator, and source-list-bound list-view Web Parts are blocked by the current deterministic profile.

## Security policy

Unique role assignments are captured as evidence. They are not replayed by the current importer. The default planning policy requires inherited permissions, so a source page with unique permissions becomes blocked unless a future reviewed security mapping is introduced.

## Existing PnP Framework reuse

The implementation composes existing PnP operations including:

- `GetPagesLibrary` for publishing-library discovery;
- `AddPublishingPage` for classic publishing-page creation;
- `GetWebParts` and `AddWebPartToWebPartPage` for shared Web Parts;
- `EnsureFolderPath` and `UploadFile` for approved dependency materialization;
- `ExecuteQueryRetry` for CSOM execution;
- `UrlUtility` and `ResourcePath` for URL/path handling.

This layer owns migration evidence, policy, approval, and verification. It should not grow local replacements for these lower-level PnP primitives.

## Current limitations

The current Enterprise Wiki profile is intentionally narrow:

- import supports create-only plans and refuses overwrite/update behavior;
- target pages must be in the root of the target publishing Pages library;
- unique permissions are captured but not restored;
- user, lookup, and taxonomy values require explicit mappings that are not implemented yet;
- only recognized field names and supported value kinds are written;
- dependency materialization occurs before page creation and there is no transaction or automatic rollback journal;
- text replacements are reviewed and digest-sealed but are currently case-insensitive text substitutions, not DOM-aware URL edits;
- list-bound and several known non-portable Web Parts are blocked rather than remapped;
- browser assertions are recorded in the plan, but browser automation is outside the library importer;
- a source fence detects capture-time mutation, but a later source change does not invalidate an already exported snapshot;
- live-tenant behavior still requires environment-specific validation in addition to unit and contract tests.

These limitations should remain explicit in plans and reports. A future implementation should add a new reviewed action or profile rather than silently relaxing a blocker.

## Validation expectations

Changes to this area should validate:

- builds for every target framework supported by `PnP.Framework`;
- JSON round-trip behavior for export, migration package, and receipt contracts;
- snapshot and plan digest mutation rejection;
- exactly one field action per captured field and one dependency action per captured dependency;
- conservative lifecycle derivation;
- package state and blocker consistency;
- Markdown report completeness;
- focused Enterprise Wiki classification, replacement, field, lifecycle, and validation tests.
