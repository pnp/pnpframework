# Publishing Page migration

`PnP.Framework.Migration.Pages.Publishing` defines the aggregate contracts and behavior that are specific to classic SharePoint Publishing Pages. It composes the reusable page capabilities documented in [../README.md](../README.md).

The first profile is Enterprise Wiki. Enterprise Wiki is an entry profile, not the owner of all participating data. Common field, reference, security, lifecycle-evidence, and classic Web Part models remain under `PnP.Framework.Migration.Pages`; publishing layout, publishing content, target lifecycle, package, report, and verification contracts belong to the Publishing family; Enterprise Wiki classification and portability policy belong to the Enterprise Wiki profile.

## Workflow and connection boundaries

```text
source connection
    -> discover and classify source page
    -> capture and seal source-only export
    -> inspect target and generate a sealed plan/report
    -> human reviews and approves the plan digest
    -> import exactly that sealed plan
    -> fresh target readback and receipt
```

| Stage | Source connection | Target connection | Writes target |
| --- | ---: | ---: | ---: |
| Discovery | Yes | No | No |
| Export | Yes | No | No |
| Plan | No | Yes | No |
| Review | No | No | No |
| Import | No | Yes | Yes |
| Verification | No | Yes | Reads after import |

Export deliberately has no target dependency. The resulting package can be stored, reviewed, and piped into planning later. Target-specific mappings and decisions start only when `EnterpriseWikiMigrationPlanner` creates a migration plan.

## Namespace ownership

| Namespace | Responsibility |
| --- | --- |
| `Pages.Publishing.Capture` | Publishing aggregate capture: common page identity plus `PublishingPageContent` and publishing layout evidence. |
| `Pages.Publishing.Layouts` | Exact Page Layout bytes, parsed controls/zones/registrations/resources, associated schema closure, target planning/admission, and create-or-exact-reuse materialization. |
| `Pages.Publishing.Lifecycle` | Conservative target Draft/Published interpretation for Publishing Pages. Source lifecycle evidence itself remains in `Pages.Lifecycle`. |
| `Pages.Publishing.Planning` | The target-specific Publishing Page plan that composes shared field, reference, replacement, and operation types. |
| `Pages.Publishing.Packaging` | Versioned Publishing Page export, migration, and receipt envelopes; canonical serialization, digests, and validation. |
| `Pages.Publishing.Execution` | Publishing-page target writes, lifecycle application, execution receipts, and failure-receipt construction. |
| `Pages.Publishing.Reporting` | Complete Markdown review output for the source snapshot and target plan. |
| `Pages.Publishing.Verification` | Publishing target probe, storage assertions, and fresh-readback verification. |
| `Pages.Publishing.EnterpriseWiki` | Enterprise Wiki discovery, classification, required layout and target policy, Web Part portability policy, export, planning, import, and file-storage facade. |

The following composed capabilities remain shared and can be used by future Wiki Page or Web Part Page families:

| Namespace | Reused capability |
| --- | --- |
| `Pages.Capture` | Capture options/status, page file probe, and source stability fence. |
| `Pages.Fields` | Full list-item field evidence, plan dispositions, and supported writes. |
| `Pages.ClassicWebParts` | Classic shared Web Part export evidence and source capture. |
| `Pages.References` | Authored reference inventory, target actions, rewriting, and payload materialization. |
| `Pages.Security` | Permission inheritance and role-assignment evidence. |
| `Pages.Lifecycle` | Source checkout, file-level, moderation, and timestamp evidence. |
| `Pages.Content` | Reviewed deterministic text replacements. |
| `Pages.Planning` | Page-wide operation and planning inputs. |

## Public Enterprise Wiki entry points

| Type | Responsibility |
| --- | --- |
| `EnterpriseWikiPageDiscovery` | Finds source Pages-library items and classifies Enterprise Wiki content types while excluding Project Pages. |
| `EnterpriseWikiPackageExporter` | Captures a source-only `PublishingPageExportPackage` and seals its snapshot digest. |
| `EnterpriseWikiMigrationPlanner` | Inspects the target, creates one explicit action per governed object, and seals a target-specific migration package and report. |
| `EnterpriseWikiMigrationImporter` | Validates the package and approved plan digest, applies the exact plan, then returns a fresh-readback receipt. |
| `EnterpriseWikiPackageFileStore` | Saves and loads export packages, migration packages, receipts, and Markdown reports, with optional validation against an `IMigrationArtifactStore`. |

These orchestration types contain Enterprise Wiki policy. Generic page readers do not switch behavior based on an Enterprise Wiki flag.

## Versioned artifacts

JSON uses camel-case property names, string enum values, explicit nulls, and case-sensitive property names. Current schema identifiers are:

| Artifact | Schema |
| --- | --- |
| Source export | `pnp-publishing-page-export/v1` |
| Target-specific migration package | `pnp-publishing-page-migration-package/v1` |
| Import receipt | `pnp-publishing-page-import-receipt/v1` |
| Nested Page Layout evidence | `pnp-publishing-page-layout/v1` |
| Nested content type schema evidence | `pnp-content-type-schema/v1` |

A breaking JSON change requires a new schema version. A CLR namespace move does not itself change JSON property names, but released CLR compatibility still needs separate consideration.

### Source export envelope

`PublishingPageExportPackage` contains:

| JSON field | Interpretation |
| --- | --- |
| `schemaVersion` | Export contract version. |
| `exportedAtUtc` | Time source capture completed. |
| `snapshot` | Complete source evidence. |
| `snapshotDigest` | SHA-256 of the canonical serialization of the complete snapshot. Any snapshot mutation invalidates it. |

`PublishingPageCaptureBundle` contains:

| JSON field | Interpretation |
| --- | --- |
| `sourceProfile` | Profile that classified the page, currently `EnterpriseWiki`. |
| `capturePolicy` | Normalized source path, whether classic Web Parts were included, and the maximum payload size per dependency. |
| `source` | Common `PageIdentity`: source web URL/path, page path, list-item ID, file ID, content type, version, length, modified time, and title. It intentionally contains no publishing layout field. |
| `layout` | Publishing-specific layout evidence: identity and gallery metadata, exact ASPX artifact, parsed server-control registrations, field-bound controls, Web Part zones, authored rendering-resource references, one evidence result per reference, and the associated content type's minimal required field-schema closure. |
| `publishingPageContent` | Complete source `PublishingPageContent` HTML. |
| `publishingPageContentSha256` | Digest of the captured publishing HTML. |
| `fields` | Every returned Pages-library field definition and typed or best-effort raw value. |
| `webParts` | Shared classic Web Part export XML and placement when capture is enabled. |
| `dependencies` | Authored references plus captured payload evidence when it can be obtained safely. |
| `security` | Permission inheritance and role-assignment evidence. |
| `lifecycle` | Source checkout type, file level, moderation status, created time, and modified time. |
| `sourceFence` | File ID, version, length, and modified time sampled before and after capture. |
| `blockers` | Source findings that make the current exact profile non-executable. |
| `warnings` | Findings requiring review that do not independently block planning. |

The source fence detects a page that changed during capture. It is not a lock, and edits made after a successful export do not alter the sealed snapshot.

### Migration package envelope

`PublishingPageMigrationPackage` embeds the exact source snapshot and adds:

| JSON field | Interpretation |
| --- | --- |
| `schemaVersion` | Migration package contract version. |
| `plannedAtUtc` | Time target analysis and plan sealing completed. |
| `exportSchemaVersion` / `exportedAtUtc` | Provenance of the embedded source export. |
| `state` | `ApprovalReady` only when the sealed plan has no blockers; otherwise `Blocked`. |
| `snapshot` | The complete source evidence used to make the decisions. |
| `plan` | Target-specific decisions and post-import assertions. |
| `snapshotDigest` | Must continue to match the embedded snapshot. |
| `planDigest` | SHA-256 over all sealed target decisions. This is the approval token supplied to Import. |
| `report` | Report metadata; the complete Markdown rendering is generated from the package. |

`PublishingPageMigrationPlan` contains:

| JSON field | Interpretation |
| --- | --- |
| `sourceSnapshotDigest` | Binds the plan to exactly one source snapshot. |
| `sourceWebUrl` / `sourcePageServerRelativeUrl` | Source boundary used by reviewed mappings. |
| `targetWebUrl` / `targetWebServerRelativeUrl` / `targetPageServerRelativeUrl` | Exact approved target web and page path. |
| `pageLayoutName` | Target publishing layout selected by the sealed layout materialization plan. |
| `operation` | Currently `CreatePage`. |
| `targetLifecycle` | Derived `Draft` or `Published` result. There is no top-level publish Boolean input. |
| `lifecycleReason` | Human-readable explanation of the derived lifecycle result. |
| `createOnly` | Currently required to be `true`; existing target pages block the plan. |
| `planningPolicy` | Normalized planning inputs copied into the sealed plan. |
| `targetProbe` | Target template, Pages library, layout, lifecycle settings, target-page existence, and dependency observations. |
| `layoutMaterialization` | Stock-reuse or deterministic digest-owned layout decision, source/target artifacts, required registrations/fields/zones, resource actions/rewrites, and associated schema plan. |
| `layoutTargetProbe` | Fresh target evidence for the selected layout path, permissions, existing bytes/association, required schema, and each owned resource path. |
| `layoutAdmission` | Typed eligibility result and issues for layout, resource, registration, permission, collision, and content-type-schema checks. |
| `fieldActions` | Exactly one reviewed decision for every captured field. |
| `dependencyActions` | Exactly one reviewed decision for every captured dependency. |
| `replacements` | Explicit source-to-target text substitutions included in the plan digest. |
| `expectedPublishingPageContentSha256` | Expected digest after approved replacements. |
| `storageAssertions` | Required storage-level readback conditions. |
| `runtimeVerification` | Typed, digest-sealed requirements for an external browser/runtime verifier. Requirements are not treated as executed by the importer. |
| `blockers` / `warnings` | Target and policy findings. `isExecutable` is derived from an empty blocker list. |

Import requires the caller's `approvedPlanDigest` to match exactly. Editing a target path, action, mapping, policy input, assertion, or lifecycle decision invalidates the package until it is replanned and reviewed again.

### Page Layout convergence model

Page Layout handling uses two deliberately different paths:

| Source evidence | Target action |
| --- | --- |
| `EnterpriseWiki.aspx` with `customizedPageStatus=1` | Reuse the reviewed target stock layout. The source stock bytes are evidence; they are not copied over the target runtime file. |
| Readable customized or non-stock ASPX | Create or exactly reuse `pnp-{safe-source-stem}-{source-layout-sha256-prefix}.aspx` in the target master page gallery. |
| Missing, denied, failed, or ambiguous layout evidence | Block. |

For a custom layout, capture parses field-bound controls, Web Part zones, server-control registrations, CSS/HTML/script references, and HTML-encoded Script Editor markup. Planning then:

1. admits only reviewed SharePoint/platform registrations;
2. builds the minimal associated content type and field closure actually required by the layout;
3. requires exact target-runtime fields or creates migration-owned fields with their source GUIDs and portable schema;
4. requires explicit term-store/term-set mappings before rebinding taxonomy field schemas;
5. copies exact source-Web or site-collection `SiteAssets`/`Style Library` resource bytes to the corresponding target owner;
6. reuses reviewed `/_layouts`, `/_controltemplates`, and known SharePoint Core Styles resources from the target runtime;
7. rewrites only the sealed authored resource strings, seals the resulting target ASPX digest, and performs create-only/exact-reuse collision checks.

Execution orders the custom layout prerequisites before the page: associated schema, rendering resources, layout ASPX, other page dependencies, then page creation. Every category receives its own mutation journal step and fresh readback.

Large ASPX/resource payloads may remain inline in JSON or live in an `IMigrationArtifactStore`. `DirectoryMigrationArtifactStore` stores them by SHA-256 under a caller-selected directory. The JSON always keeps digest, length, media type, original name, availability, and optional lineage; package load/import can validate the supplied store before any target mutation.

### Markdown review report coverage

`PublishingPageMigrationReportBuilder` is the human review projection of the authoritative JSON package. It exposes every source and target decision without dumping unbounded payloads directly into Markdown:

- package identity, schema versions, timestamps, state, snapshot digest, and approval plan digest;
- source page identity, capture policy, stability fence, lifecycle evidence, publishing HTML digest, and derived lifecycle;
- every Page Layout scalar property, exact artifact metadata, registration, parsed control, field binding, zone, resource reference, resource evidence state, byte digest, source lineage, and diagnostic;
- every associated content type and field-schema property, field link, portable digest, ownership/role, taxonomy binding, and source diagnostic;
- every captured page-item field and its complete recovery representation plus exactly one planned field action;
- every classic Web Part export and every authored page dependency plus its target action;
- the layout/schema/resource materialization plan, all resource rewrites, target probes, typed admission issues, and approved taxonomy schema mappings;
- target page/library evidence, text replacements, storage assertions, runtime-verification requirements, blockers, and warnings.

Large HTML, XML, JSON, ASPX, and Base64 values are represented as length, SHA-256, and a bounded preview. Their complete value remains in the JSON package or artifact store. A reviewer should treat `snapshotDigest` as source-evidence identity and `planDigest` as the complete approval boundary.

### Target probe

`PublishingPageTargetSnapshot` records the facts used during planning:

| JSON field | Interpretation |
| --- | --- |
| `webUrl` / `webServerRelativeUrl` | Resolved target web identity. |
| `webTemplate` / `webConfiguration` | Target site template evidence used by the profile. |
| `pagesLibraryServerRelativeUrl` / `pagesLibraryBaseTemplate` | Resolved target Pages library. |
| `enableVersioning` / `enableMinorVersions` / `enableModeration` / `forceCheckout` / `draftVersionVisibility` | Target library lifecycle behavior. |
| `pageContentTypeId` | Target publishing page content type selected by the profile. |
| `pageLayoutUrl` / `pageLayoutExists` | Compatibility summary of the layout selected by the profile. Detailed exact-byte and schema/resource evidence lives in `layoutTargetProbe`. |
| `targetPageExists` | Create-only collision check. |
| `existingDependencyPaths` | Dependency targets already present when the plan was created. |

Import rechecks critical target facts before writing so that a stale plan does not silently execute against materially changed target state.

### Import receipt

`PublishingPageImportReceipt` contains:

| JSON field | Interpretation |
| --- | --- |
| `schemaVersion` | Receipt contract version. |
| `startedAtUtc` / `completedAtUtc` | Import execution interval. |
| `operationId` / `executionStatus` | Independent attempt identity and `NotStarted`, `Running`, `Succeeded`, or `FailedUnexpectedly` outcome. |
| `admissionFailure` / `mutationStarted` / `steps` | Zero-mutation admission result or the ordered write-ahead mutation receipts. |
| `approvedPlanDigest` | Approval token actually presented to Import. |
| `targetWebUrl` / `targetPageServerRelativeUrl` | Executed target. |
| `targetFileUniqueId` / `targetListItemId` / `targetContentTypeId` / `targetVersionLabel` | Persisted target identity returned by fresh readback. |
| `expectedLifecycle` | Lifecycle sealed in the plan. |
| `actualFileLevel` / `actualCheckOutType` / `actualModerationStatus` | Fresh lifecycle evidence. |
| `lifecycleMatched` | Whether persisted evidence satisfies the planned lifecycle. |
| `expectedPublishingPageContentSha256` / `persistedPublishingPageContentSha256` | Expected and read-back content digests. |
| `storageContentEqual` | Whether storage-level content matches. |
| `importedWebPartCount` / `materializedDependencyCount` | Applied object counts. |
| `webPartsMatched` / `webPartResults` | Per-Web-Part export digest, zone, order, and hidden-state readback results. |
| `fieldResults` | Per-field write result and diagnostics. |
| `freshReadbackPassed` | Whether required fresh-readback assertions passed. |
| `storageVerificationStatus` | `Passed` only when content digest, Web Parts, fields, content type, and lifecycle all match. |
| `runtimeVerificationStatus` / `acceptanceStatus` | Runtime work remains `Pending` until an external verifier supplies evidence; storage success alone is not final acceptance when runtime requirements exist. |
| `warnings` | Non-fatal import/verification findings. |

The receipt records observed outcome and ordered mutation steps. It is not a promise of a cross-object transaction or automatic global rollback.

## Complete field capture, selective restore

Field capture and field restore intentionally have different scopes. Every field returned for the source Pages-library item is retained with:

- field ID, internal name, title, type, schema XML, and read-only/hidden/required flags;
- a typed representation for supported scalar, URL, lookup, taxonomy, multi-value, and binary values;
- best-effort runtime type, text, JSON, or Base64 evidence when no supported representation exists;
- capture status and diagnostics.

Planning creates exactly one `PageFieldAction` per captured field:

| Disposition | Interpretation |
| --- | --- |
| `Apply` | Recognized, non-empty, writable, target-present, type-compatible, and supported; Import writes it. |
| `AlreadyHandled` | Page creation, content, or layout logic owns the property. |
| `SkipEmpty` | No source value needs restoring. |
| `SkipReadOnly` | SharePoint owns the source or target field. |
| `SkipCalculated` | SharePoint recomputes the value. |
| `TargetFieldMissing` | A recognized source field is absent at the target. |
| `TargetTypeMismatch` | Source and target field types differ. |
| `RequiresMapping` | User, lookup, or taxonomy identity cannot be copied safely across sites without an explicit mapping. |
| `EvidenceOnly` | The snapshot retains complete evidence, but the current importer does not own restoration. |
| `CaptureUnavailable` | The definition was captured, but no restorable value was returned. |
| `Block` | The exact plan cannot execute. |

The Enterprise Wiki profile currently recognizes a reviewed subset of publishing metadata. Unknown fields are never discarded and are never guessed into a target field. This preserves a recovery snapshot for a later mapper without weakening current import safety. Page Layout field-schema closure is a separate concern: it recreates only fields proven necessary to render the approved layout and does not imply that arbitrary source page-item values are replayed.

## Publishing lifecycle policy

Target lifecycle is family-specific because Draft/Published semantics depend on publishing storage and library behavior. Source facts remain shared in `PageLifecycleSnapshot`.

The current policy returns `Published` only when all available evidence is unambiguous:

- source file level is `Published`;
- checkout type is `None`;
- moderation is absent or approved (`0`).

Every other, missing, or contradictory state becomes `Draft`.

| Source evidence | Target result | Interpretation |
| --- | --- | --- |
| `level=Draft`, `checkOutType=Online`, `moderationStatus=3` | `Draft` | Source is not an unconflicted published version. |
| `level=Published`, `checkOutType=None`, `moderationStatus=0` | `Published` | All captured evidence agrees that the source is published. |
| missing lifecycle evidence | `Draft` | Conservative default. |

There is no caller-provided top-level `publish` switch. If a planned field write fails, Import avoids publishing and records a warning.

## References and classic Web Parts

Authored references are inventoried separately so each receives an explicit action:

| Disposition | Behavior |
| --- | --- |
| `PreserveExternal` | Leave an allowed external reference unchanged. |
| `RewriteToTarget` | Rewrite a same-tenant or same-web reference to the reviewed target location. |
| `MaterializeAtTarget` | Upload the captured same-web payload before page creation. |
| `Delegate` | Reserve handling for another reviewed migration owner. |
| `Block` | Stop Import until the unsupported reference is resolved and the package is replanned. |

Same-tenant iframes, resources outside the captured web boundary, and missing restorable payloads block the exact profile. The default per-dependency capture limit is 10 MiB.

`ClassicWebPartSnapshotReader` only captures common evidence: export XML, ID, title, zone, index, hidden state, and digest. It does not decide portability. `EnterpriseWikiWebPartPolicy` separately blocks known unsupported types such as RSS Aggregator and source-list-bound list-view Web Parts. This lets another page profile reuse the snapshot reader with a different portability policy.

## Security policy

Unique role assignments are captured as common evidence. The current Enterprise Wiki importer does not replay them. With the default `RequireInheritedPermissions` planning option, unique source permissions become a blocker until a reviewed cross-site security mapping exists.

## Existing PnP Framework reuse

The implementation composes established PnP Framework operations, including:

- `GetPagesLibrary` for publishing-library discovery;
- `AddPublishingPage` for classic publishing-page creation;
- `GetWebParts` and `AddWebPartToWebPartPage` for shared classic Web Parts;
- `EnsureFolderPath` and `UploadFile` for approved dependency materialization;
- `EnsureSiteAssetsLibrary` plus list/root-folder checks for Page Layout resource ownership;
- `ExecuteQueryRetry` for CSOM execution;
- `UrlUtility`, `ResourcePath`, and existing page/file extensions for URL and storage handling.

This layer owns migration evidence, policy, approval, and verification. It should not duplicate lower-level CSOM plumbing already provided by PnP Framework.

## Current limitations

The current Enterprise Wiki profile is intentionally narrow:

- only create-only plans are executable; overwrite/update is refused;
- target pages must be in the root of the target Publishing Pages library;
- unique permissions are captured but not restored;
- user, lookup, and page-item taxonomy values still require explicit identity/value mappings; Page Layout taxonomy field *schema* supports reviewed store/set rebinding, but term/value migration is separate;
- only recognized fields with supported values and compatible target definitions are written;
- dependency materialization happens before page creation; mutation intents and receipts are journaled, but there is no cross-object transaction or automatic global rollback;
- replacements are reviewed and digest-sealed but are case-insensitive text substitutions rather than DOM-aware URL edits;
- source-list-bound and known non-portable Web Parts are still blocked instead of remapped;
- typed runtime verification requirements are sealed in the plan, but browser automation and its evidence receipt remain outside the library importer;
- a source fence detects capture-time mutation but does not invalidate an export after a later source edit;
- live-tenant behavior still requires environment-specific validation in addition to unit and contract tests.

A future implementation should add an explicit action or a narrower/new profile instead of silently relaxing one of these blockers.

## Validation expectations

Changes to this family should validate:

- all target frameworks supported by `PnP.Framework` build;
- export, migration package, and receipt JSON round trips;
- snapshot and plan mutation invalidates the corresponding digest;
- exactly one field action exists per captured field and one dependency action per captured dependency;
- shared `Pages` code has no dependency on `Publishing` or `EnterpriseWiki`;
- publishing lifecycle derivation remains conservative;
- package state agrees with blocker state;
- the Markdown report exposes complete source evidence and every plan action;
- focused classification, replacement, field, lifecycle, Web Part policy, and validation tests pass.
