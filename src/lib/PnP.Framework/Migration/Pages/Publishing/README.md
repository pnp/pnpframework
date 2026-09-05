# Publishing Page migration

`PnP.Framework.Migration.Pages.Publishing` defines the aggregate contracts and behavior that are specific to classic SharePoint Publishing Pages. It composes the reusable page capabilities documented in [../README.md](../README.md).

The first workflow is Enterprise Wiki v1. Enterprise Wiki is an entry facade and validation cohort, not the owner of all participating data. Common ASPX, runtime, profile-signal, ingredient, field, reference, security, lifecycle-evidence, and classic Web Part models remain under `PnP.Framework.Migration.Pages`; publishing layout, content, target lifecycle, planning, package, report, execution, and verification belong to the Publishing family.

For the cross-domain artifact chain and execution semantics, see the Migration design documents for the [package model](../../docs/package-model.md), [page classification and ingredient policy](../../docs/page-classification-and-ingredient-policy.md), [object lifecycle](../../docs/object-lifecycle.md), and [execution and verification](../../docs/execution-and-verification.md).

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
| `Pages.Publishing.Capture` | Publishing aggregate capture and generic package export: common page identity/artifact/runtime/profile/ingredient evidence plus `PublishingPageContent` and layout evidence. |
| `Pages.Publishing.Layouts` | Exact Page Layout bytes, parsed controls/zones/registrations/resources, associated schema closure, target planning/admission, and create-or-exact-reuse materialization. |
| `Pages.Publishing.Lifecycle` | Conservative target Draft/Published interpretation for Publishing Pages. Source lifecycle evidence itself remains in `Pages.Lifecycle`. |
| `Pages.Publishing.Profiles` | Publishing workflow policy and EW-v1 cohort/profile projection. |
| `Pages.Publishing.Ingredients` | Thin graph/action coordinators plus object-owned Core, Layout, Topology, List-schema, List-content, Web Part, and Reference projectors. |
| `Pages.Publishing.Planning` | Generic target-specific orchestration plus separate lifecycle/field policy and topology/List/Web Part dependency planning. |
| `Pages.Publishing.Packaging` | Versioned Publishing Page export, migration, and receipt envelopes; canonical serialization, digests, and validation. |
| `Pages.Publishing.Execution` | Publishing-page target writes, lifecycle application, execution receipts, and failure-receipt construction. |
| `Pages.Publishing.Reporting` | Complete Markdown review output for the source snapshot and target plan. |
| `Pages.Publishing.Verification` | Publishing target probe, storage assertions, and fresh-readback verification. |
| `Pages.Publishing.EnterpriseWiki` | Thin public Enterprise Wiki v1 export/planning/import facades, discovery, and EW-named file storage. |

The following composed capabilities remain shared and can be used by future Wiki Page or Web Part Page families:

| Namespace | Reused capability |
| --- | --- |
| `Pages.Capture` | Capture options/status, page file probe, and source stability fence. |
| `Pages.Markup` | Exact source ASPX bytes and parsed Page directive. |
| `Pages.Runtime` | CLR-first runtime-adapter resolution. |
| `Pages.Profiles` / `Pages.Cohorts` | Non-exclusive product signals and versioned validation-population decisions. |
| `Pages.Ingredients` | Canonical dependency graph, semantic dispositions, and aggregate outcome evaluation. |
| `Pages.Fields` | Full list-item field evidence, plan dispositions, and supported writes. |
| `Pages.ClassicWebParts` | Classic shared Web Part export evidence and source capture. |
| `Pages.References` | Authored reference inventory, target actions, rewriting, and payload materialization. |
| `Pages.Security` | Permission inheritance and role-assignment evidence. |
| `Pages.Lifecycle` | Source checkout, file-level, moderation, and timestamp evidence. |
| `Pages.Content` | Reviewed deterministic text replacements. |
| `Pages.Planning` | Page-wide operation and planning inputs. |
| `Migration.Topology` | SPSite/SPWeb ancestor closure, target mapping, owned child-Web creation/recovery, and runtime Web-ID receipts. |
| `Migration.Schema.ContentTypes` | Custom site-content-type parent/field closure used by Page Layouts and Lists. |
| `Migration.Lists` | List/library, List content type, field, View, current item/file/attachment, and lookup-ID closure. |

## Public Enterprise Wiki entry points

| Type | Responsibility |
| --- | --- |
| `EnterpriseWikiPageDiscovery` | Finds source Pages-library items and classifies Enterprise Wiki content types while excluding Project Pages. |
| `EnterpriseWikiPackageExporter` | Captures a source-only `PublishingPageExportPackage` and seals its snapshot digest. |
| `EnterpriseWikiMigrationPlanner` | Inspects the target, creates one explicit action per governed object, and seals a target-specific migration package and report. |
| `EnterpriseWikiMigrationImporter` | Validates the package, admits the approved plan against fresh target state, applies the exact plan, then returns a fresh-readback receipt. Invalid package contracts currently throw before a receipt is created. |
| `EnterpriseWikiPackageFileStore` | Saves and loads export packages, migration packages, receipts, and Markdown reports, with optional validation against an `IMigrationArtifactStore`. |

These public types select `EnterpriseWikiV1WorkflowPolicy` and delegate to Publishing-family services. Generic readers and planners do not switch behavior based on an Enterprise Wiki flag.

## Versioned artifacts

JSON uses camel-case property names, string enum values, explicit nulls, and case-sensitive property names. Current schema identifiers are:

| Artifact | Schema |
| --- | --- |
| Source export | `pnp-publishing-page-export/v2` |
| Target-specific migration package | `pnp-publishing-page-migration-package/v2` |
| Import receipt | `pnp-publishing-page-import-receipt/v2` |
| Nested source ASPX artifact | `pnp-page-artifact/v1` |
| Nested page runtime resolution | `pnp-page-runtime/v1` |
| Nested canonical ingredient graph | `pnp-page-ingredient-graph/v1` |
| Nested Page Layout evidence | `pnp-publishing-page-layout/v1` |
| Nested content type schema evidence | `pnp-content-type-schema/v1` |
| Nested source topology | `pnp-source-topology/v1` |
| Nested topology plan / target analysis | `pnp-topology-plan/v1` / `pnp-topology-target-analysis/v1` |
| Nested List dependency / List plan | `pnp-list-dependency/v1` / `pnp-list-migration-plan/v1` |

A breaking JSON change requires a new schema version. A CLR namespace move does not itself change JSON property names, but released CLR compatibility still needs separate consideration.

### Source export envelope

`PublishingPageExportPackage` contains:

| JSON field | Interpretation |
| --- | --- |
| `schemaVersion` | Export contract version. |
| `exportedAtUtc` | Time source capture completed. |
| `selection` | Workflow ID and versioned validation-cohort assessment. |
| `selectionDigest` | SHA-256 over the exact workflow/cohort selection. |
| `snapshot` | Complete source evidence. |
| `snapshotDigest` | SHA-256 of the canonical serialization of the complete snapshot. Any snapshot mutation invalidates it. |

`PublishingPageCaptureBundle` contains:

| JSON field | Interpretation |
| --- | --- |
| `capturePolicy` | Normalized source path, whether classic Web Parts were included, and the maximum payload size per dependency. |
| `source` | Common `PageIdentity`: source web URL/path, page path, list-item ID, file ID, content type, version, length, modified time, and title. It intentionally contains no publishing layout field. |
| `pageArtifact` | Exact source ASPX artifact plus parsed Page directive, availability, and diagnostics. |
| `runtime` | Adapter selected from CLR type evidence, with Content Type fallback made explicit. |
| `profileSignals` | Non-exclusive Content Type/layout/field trait signals; multiple profiles can apply. |
| `ingredientGraph` | Canonical page/external-content nodes and required/conditional/optional dependency edges. Nodes cover owner Webs, layout resources and associated fields, List/site schema, every current item/document/attachment/View, Web Parts, and references in addition to the page core. |
| `layout` | Publishing-specific layout evidence: identity and gallery metadata, exact ASPX artifact, parsed server-control registrations, field-bound controls, Web Part zones, authored rendering-resource references, one evidence result per reference, and the associated content type's minimal required field-schema closure. |
| `publishingPageContent` | Complete source `PublishingPageContent` HTML. |
| `publishingPageContentSha256` | Digest of the captured publishing HTML. |
| `fields` | Every returned Pages-library field definition and typed or best-effort raw value. |
| `webParts` | Shared classic Web Part export XML and placement when capture is enabled. |
| `listWebPartBindings` | Parsed list-bound Web Part identity: page/list owner Webs, source List/View IDs, path, TitleUrl, XML definition, JSLink/XslLink, and exact source export. |
| `listDependencies` | Complete current evidence for every required List/library: settings, fields, List/site content types, explicit CT order, Views, items, folders, exact file bytes, and attachments. |
| `listLookupDependencies` | Directed lookup edges used to build the target materialization order and detect cycles before mutation. |
| `sourceTopology` | Source SPSite plus the complete SPWeb ancestor closure needed to preserve each object's ownership boundary. |
| `dependencies` | Authored references plus captured payload evidence when it can be obtained safely. |
| `security` | Permission inheritance and role-assignment evidence. |
| `lifecycle` | Source checkout type, file level, moderation status, created time, and modified time. |
| `sourceFence` | File ID, version, length, and modified time sampled before and after capture. |
| `blockers` | Source findings that make the selected workflow non-executable. |
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
| `selection` | Exact workflow and cohort assessment copied from the source export. |
| `selectionDigest` | Must match both the embedded selection and the assessment recomputed by the selected workflow policy. |
| `snapshot` | The complete source evidence used to make the decisions. |
| `plan` | Target-specific decisions and post-import assertions. |
| `snapshotDigest` | Must continue to match the embedded snapshot. |
| `planDigest` | SHA-256 over the complete sealed plan, including policy, probes, actions, issues, and assertions. This is the approval token supplied to Import. |
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
| `topology` | Parent-preserving source Site/Web to target Site/Web map, deterministic target shape, provenance identities, and semantic digest. |
| `topologyTargetAnalysis` | Read-only target collision/ownership analysis for each mapped Site and Web. Missing child Webs can be planned for creation; existing unowned paths block. |
| `listMigration` | Lookup-ordered per-List plans, field/View/site-CT actions, target paths/titles, ownership digests, and target probes. |
| `webPartActions` | One action per shared Web Part. List-bound parts are replayed only after target Web/List/View IDs are available and their XML can be rebound. |
| `replacements` | Explicit source-to-target text substitutions included in the plan digest. |
| `expectedPublishingPageContentSha256` | Expected digest after approved replacements. |
| `storageAssertions` | Required storage-level readback conditions. |
| `runtimeVerification` | Typed, digest-sealed requirements for an external browser/runtime verifier. Requirements are not treated as executed by the importer. |
| `ingredientActions` | One capability/disposition/realization/target/policy/verification record for every non-empty canonical ingredient. A dependency release is accepted only from a `Transform` and only for one of that ingredient's real required edges. |
| `migrationOutcome` | Aggregate `Exact`, `ExecutableWithTransform`, `ExecutableWithLoss`, `Blocked`, or `Unknown` result. |
| `ingredientIssues` | Recomputed missing-action and required-dependency-closure issues. |
| `blockers` / `warnings` | Target and policy findings. `isExecutable` requires no blockers and an executable ingredient outcome. |

Import requires the caller's `approvedPlanDigest` to match exactly. Editing any sealed plan content, including a target path, planning probe, action, mapping, policy input, issue, assertion, or lifecycle decision, invalidates the package until it is replanned and reviewed again.

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
- the complete source SPSite/SPWeb hierarchy, every approved mapping, observed target identity/shape/provenance, and create/reuse/recover/block disposition;
- every List setting, field schema, custom site-content-type ancestor and field closure, List-local content type and FieldLink, explicit CT order, View, current item value, folder/file artifact, attachment, lookup edge, target action, target probe, and typed issue;
- every item value's typed form plus raw runtime type/text/JSON recovery evidence, including values the current importer will not write;
- every canonical ingredient node/edge and its independent action, including explicitly dropped unused List fields and the retained snapshot that makes later recovery possible;
- every classic Web Part export, parsed List binding, approved copy/rebind/block action, and every authored page dependency plus its target action;
- the layout/schema/resource materialization plan, all resource rewrites, target probes, typed admission issues, and approved taxonomy schema mappings;
- target page/library evidence, text replacements, storage assertions, runtime-verification requirements, blockers, and warnings.

Large HTML, XML, JSON, ASPX, and Base64 values are represented as length, SHA-256, and a bounded preview. Their complete value remains in the JSON package or artifact store. A reviewer should treat `snapshotDigest` as source-evidence identity and `planDigest` as the complete approval boundary.

### Target probe

`PublishingPageTargetSnapshot` records the facts used during planning:

| JSON field | Interpretation |
| --- | --- |
| `webUrl` / `webServerRelativeUrl` | Resolved target web identity. |
| `webTemplate` / `webConfiguration` | Target site template evidence used by planning. |
| `pagesLibraryServerRelativeUrl` / `pagesLibraryBaseTemplate` | Resolved target Pages library. |
| `enableVersioning` / `enableMinorVersions` / `enableModeration` / `forceCheckout` / `draftVersionVisibility` | Target library lifecycle behavior. |
| `pageContentTypeId` | One exact Pages-library Content Type ID derived from the approved Page Layout association. Ambiguous descendants block planning, and import verifies exact equality. |
| `pageLayoutUrl` / `pageLayoutExists` | Compatibility summary of the approved layout. Detailed exact-byte and schema/resource evidence lives in `layoutTargetProbe`. |
| `targetPageExists` | Create-only collision check. |
| `existingDependencyPaths` | Dependency targets already present when the plan was created. |

Import rechecks critical target facts before writing so that a stale plan does not silently execute against materially changed target state. The shared List target analyzer re-probes every List and custom site content type whose owner Web already exists; objects below a still-missing, approved child Web remain explicitly deferred and are probed immediately after topology materialization.

### Import receipt

`PublishingPageImportReceipt` contains:

| JSON field | Interpretation |
| --- | --- |
| `schemaVersion` | Receipt contract version. |
| `startedAtUtc` / `completedAtUtc` | Import execution interval. |
| `operationId` / `executionStatus` | Independent attempt identity and `NotStarted`, `Running`, `Succeeded`, or `FailedUnexpectedly` outcome. |
| `admissionFailure` / `mutationStarted` / `steps` | Zero-mutation approval/target-admission result or the ordered write-ahead mutation receipts. Malformed contracts throw before this receipt exists. |
| `approvedPlanDigest` | Approval token presented to an admitted execution. Current rejection receipts record the package digest rather than retaining a mismatched caller candidate. |
| `targetWebUrl` / `targetPageServerRelativeUrl` | Executed target. |
| `targetFileUniqueId` / `targetListItemId` / `targetContentTypeId` / `targetVersionLabel` | Persisted target identity returned by fresh readback. |
| `expectedLifecycle` | Lifecycle sealed in the plan. |
| `actualFileLevel` / `actualCheckOutType` / `actualModerationStatus` | Fresh lifecycle evidence. |
| `lifecycleMatched` | Whether persisted evidence satisfies the planned lifecycle. |
| `expectedPublishingPageContentSha256` / `persistedPublishingPageContentSha256` | Expected and read-back content digests. |
| `storageContentEqual` | Whether storage-level content matches. |
| `importedWebPartCount` / `materializedDependencyCount` | Applied object counts. |
| `webPartsMatched` / `webPartResults` | Per-Web-Part export digest, zone, order, and hidden-state readback results. |
| `topologyMaterialization` / `topologyMatched` | Source-to-target Site/Web runtime IDs, execution dispositions, mapping digests, and final whole-topology readback status. |
| `listMaterializations` / `listsMatched` | Per-List target Web/List IDs, item/View/content-type maps, verified object counts, diagnostics, and final whole-closure readback status. |
| `fieldResults` | Per-field write result and diagnostics, including target-local taxonomy WssId materialization receipts. |
| `taxonomyRelationshipsMatched` / `taxonomyRelationshipResults` | Per-executed-value fresh verification of exact Term state, page value, hidden-list identity, and `TaxCatchAll`; evidence-only relationships make no target claim and are omitted. |
| `freshReadbackPassed` | Whether required fresh-readback assertions passed. |
| `storageVerificationStatus` | `Passed` only when required topology, List, content, Web Part, field, content-type, and lifecycle readback all match. |
| `runtimeVerificationStatus` / `acceptanceStatus` | Runtime work remains `Pending` when runtime requirements exist. The runtime receipt contract is defined, but the library does not yet reconcile it into a new final acceptance record. |
| `warnings` | Non-fatal import/verification findings. |

The receipt records observed outcome and ordered mutation steps. It is not a promise of a cross-object transaction or automatic global rollback.

## Cross-site topology and List dependency closure

A list-bound Publishing Page is not portable by copying its ASPX and Web Part XML alone. The Web Part's source `WebId`, `ListId`/`ListName`, View ID, lookup item IDs, and site-local taxonomy WssIds are runtime identities. The package therefore treats them as a dependency graph:

```text
source Site/Web ancestor closure
    -> approved target Web map
    -> Page Layout schema/resources/layout and other dependency artifacts
    -> lookup Lists before consuming Lists
    -> required site and List-local content types and fields
    -> current folders/items/files/attachments
    -> public and page-bound Views
    -> source-to-target Web/List/View/item ID catalog
    -> create the page
    -> transformed content/fields and rebound classic Web Parts
    -> derived page lifecycle
    -> final topology/List/page fresh readback
```

The current topology executor requires an existing target site collection. Its root and the explicitly connected page Web are approved hosts; missing mapped child Webs can be created with the sealed target template, and an interrupted child-Web creation can be claimed only when its exact recovery description, title, template, parent, and path agree. Every reusable created Web must then expose both the source-qualified original identifier and the exact mapping digest.

List creation follows the same ownership model. A free path and title produce `CreateOwned`; an existing path is reusable only when template, target title, original identifier, and semantic plan digest all agree. A same-title List at another path is reported separately because SharePoint can reject the new List even though its path is free.

### List execution and verification boundary

For each List in lookup topological order, Import performs:

1. materialize the de-duplicated custom site-content-type parent closure;
2. create or exactly reuse the target List;
3. add every required site content type and record source List-CT ID to target List-CT ID mappings;
4. create/require fields, with lookup schema rebound to target Web/List IDs and taxonomy schema rebound only through an approved store/set mapping;
5. apply captured List content-type metadata, FieldLink flags, and explicit New-button content-type order;
6. materialize folders, current items/files, and attachments, then write lookup values through the target item-ID catalogs;
7. create/reuse supported public and page-bound Views;
8. freshly read the entire supported closure again before the page receipt can pass.

Fresh List verification checks target identity and settings, ownership markers, planned fields and portable schema, List content-type parent/metadata/FieldLinks/order, supported View state, source-to-target item mappings, every written current value, taxonomy Term GUID/label rather than WssId, exact current document and attachment bytes, and object counts. A List mismatch makes `listsMatched=false`, which makes storage acceptance fail even if page creation itself succeeded.

### Complete report field guide with examples

The JSON package is authoritative; the Markdown report is generated from it and prints a bounded representation of every reviewable property. The following examples use representative SharePoint values and deliberately omit tenant/customer names.

| Report/JSON field | Example | Interpretation |
| --- | --- | --- |
| `snapshot.sourceTopology.siteId` | `11111111-1111-1111-1111-111111111111` | Source SPSite evidence. It identifies the mapping input; it is never assigned to the target. |
| `snapshot.sourceTopology.webs[].parentWebId` | root Web GUID | Proves a source child Web's parent. The target child must be under the mapped target parent, not flattened under whichever Web hosts the page command. |
| `plan.topology.siteCollections[].targetMode` | `ExistingTargetSite` | The current importer probes an existing target site collection. `CreateTargetSite` remains blocked until a tenant-scoped executor exists. |
| `plan.topologyTargetAnalysis...disposition` | `CreateOwned` | The target child-Web path is free and the sealed plan may create it. After creation, fresh analysis must return `ReuseOwned`. |
| `snapshot.listDependencies[].baseTemplate` | `101` | Document library template. Templates outside the reviewed set block before mutation. |
| `snapshot.listDependencies[].hasExplicitUniqueContentTypeOrder` | `true` | SharePoint returned an explicit New-button order. `false` is distinct: a null order means all allowed content types are visible. |
| `snapshot.listDependencies[].uniqueContentTypeOrder[]` | source List-local custom Document CT ID | The source order is captured using source IDs, then translated to target-generated List CT IDs. Folder/UntypedDocument children are filtered because SharePoint rejects them in this property. |
| `snapshot.listDependencies[].fields[].portableSchemaSha256` | 64-character SHA-256 | Schema equality after removing non-portable storage slots and runtime identities such as `List`, `WebId`, `SourceID`, `Version`, `ColName`, and `RowOrdinal`. |
| `plan.listMigration.lists[].fields[].disposition` | `MapLookup` | Create/reuse this field only after its dependency List exists; rewrite its schema to target Web/List GUIDs and translate every source lookup item ID. |
| `plan.listMigration.lists[].fields[].disposition` | `RequireTargetRuntime` for `Modified` | The target template must supply the field, but SharePoint-owned/read-only audit values are intentionally not written. |
| `snapshot.listDependencies[].siteContentTypes[].contentTypeId` | custom child of `0x0101` | A custom Document content type remains custom. Runtime classification uses exact known IDs, not the broad `0x0101` prefix. |
| `...requiredFieldClosure[].taxonomy.sourceTermSetId` | source TermSet GUID | Source taxonomy schema identity. Planning requires an explicit target store/set mapping; source WssId never participates. |
| `snapshot.listDependencies[].contentTypes[].id` | source List-local CT ID | Evidence identity. Its exact site parent resolves the target List CT even when the source name was customized; metadata and known FieldLink flags are replayed, while the generated target child ID is stored in `receipt.listMaterializations[].targetContentTypeIds`. |
| `snapshot.listDependencies[].views[].listViewXmlSha256` | 64-character SHA-256 | Digest of the complete captured View XML. Supported query/fields/paging/JSLink state is applied now; custom JSLink/XslLink resources and full XML fidelity remain explicit gaps. |
| `snapshot.listDependencies[].items[].values[].kind` | `Unsupported` | The current serializer did not recognize the runtime object. `rawType`, `rawValue`, and `rawValueJson` retain recovery evidence; planning will not guess a write. |
| `snapshot.listDependencies[].items[].values[].taxonomyValues[].wssId` | `269` | Source site-collection cache-row evidence only. Import writes the mapped Term through the target taxonomy field and verifies target Term GUID/label with a target-allocated WssId. |
| `snapshot.fields[].taxonomyBinding.boundTermSetId` | source Wiki Categories TermSet GUID | The field's source binding. It is mapped explicitly and is not changed merely to make an invalid value valid. |
| `snapshot.fields[].taxonomyValues[].relationship.state` | `LiveOutsideBoundTermSet` | The GUID is live, but not in the field's bound TermSet. This is source data to reproduce, not an instruction to move or recreate the Term. |
| `snapshot.fields[].taxonomyValues[].relationship.valueHiddenListEntry` | WssId plus store/set/term, localized label/path and CatchAll data | Exact source site-collection cache evidence used to allocate and verify a target-local WssId. |
| `snapshot.fields[].taxonomyValues[].relationship.evidenceSha256` | 64-character SHA-256 | Binds the value relationship to the complete field value set, field binding, live resolution, hidden rows, timestamp, and diagnostics. |
| `plan.taxonomyRelationshipActions[].disposition` | `PreserveDanglingTermAbsent` | Keep the Term absent and reproduce the invalid relationship. A newly live target Term with that GUID makes fresh admission fail. |
| `plan.taxonomyRelationshipActions[].disposition` | `RetainEvidenceOnly` | The owning field is not selected for replay. The exact relationship proof stays sealed for later recovery, but this plan does not probe, materialize, or verify a target taxonomy relationship. |
| `receipt.taxonomyRelationshipResults[].relationshipStateMatched` | `true` | A fresh context confirmed that the target relationship is still live-in-bound, live-outside-bound, or dangling exactly as approved. |
| `snapshot.listDependencies[].items[].document.content.artifact.sha256` | file byte digest | Exact current document bytes. The target file must read back with the same length and digest; version history is not represented by this field. |
| `plan.listMigration.lists[].targetProbe.sameTitleDifferentPaths[]` | `/sites/target/Shared Documents` | A real SharePoint collision shape: the desired `/Documents` path can be free while a template-created library already owns title `Documents`. Strict mode blocks rather than renaming an unrelated List. |
| `snapshot.listWebPartBindings[].sourceViewId` | source View GUID | Captured binding input. It is replaced with `targetViewIds[sourceViewId]` only after the target View passes readback. |
| `receipt.listMaterializations[].targetItemIds` | `42 -> 7` | Source item 42 became target item 7. Lookup consumers use 7; they never attempt to preserve 42. |
| `receipt.listMaterializations[].disposition` | `CreateOwned` or `ReuseOwned` | Actual disposition observed by the execution-time fresh preflight, which may safely advance from the planning-time state when another executor completed the same identity and digest. |
| `receipt.listMaterializations[].freshReadbackPassed` | `true` | All currently owned List assertions passed after mutation. Counter fields report how many fields, CTs, Views, items, files/folders, and attachments were inspected. |

Large XML, HTML, JSON, and Base64 cells show length, SHA-256, and a bounded preview. This is not truncation of the package: the complete value remains in JSON or its content-addressed artifact store. A report action answers “what this importer will do now”; raw evidence answers “what could a later importer recover.”

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
| `ApplyTaxonomyRelationships` | Every taxonomy value has a separately reviewed executable relationship action. Import reproduces those exact relationships without creating or substituting Terms. |
| `AlreadyHandled` | Page creation, content, or layout logic owns the property. |
| `SkipEmpty` | No source value needs restoring. |
| `SkipReadOnly` | SharePoint owns the source or target field. |
| `SkipCalculated` | SharePoint recomputes the value. |
| `TargetRuntime` | Source schema is SharePoint-owned and the target exposes an equivalent same-name, same-type field; the target runtime regenerates the value. |
| `TargetFieldMissing` | A recognized source field is absent at the target. |
| `TargetTypeMismatch` | Source and target field types differ. |
| `RequiresMapping` | User or lookup identity cannot be copied safely across sites without an explicit mapping. Taxonomy uses typed per-value relationship actions and blocks when they are incomplete. |
| `EvidenceOnly` | The snapshot retains complete evidence, but the current importer does not own restoration. |
| `CaptureUnavailable` | The definition was captured, but no restorable value was returned. |
| `Block` | The exact plan cannot execute. |

At ingredient level, a field action with no source value and no required material projects as `Drop`, not `Block`. The complete field definition and capture diagnostics remain in the snapshot. If the ingredient is required by the layout or another retained node, normal dependency-closure validation still prevents an unsafe drop.

The Enterprise Wiki v1 workflow currently recognizes a reviewed subset of publishing metadata. Unknown field evidence is never discarded from the snapshot and is never guessed into a target field; an empty action may still be explicitly dropped from the execution closure. This preserves a recovery snapshot for a later mapper without weakening current import safety. Page Layout field-schema closure is a separate concern: it recreates only fields proven necessary to render the approved layout and does not imply that arbitrary source page-item values are replayed.

Taxonomy follows the same capture-wide/restore-narrow rule at value granularity. Every captured taxonomy value has one relationship action. Values owned by a selected field receive strict target-aware replay or block actions; values in an unselected field receive `RetainEvidenceOnly`, project as delegated evidence, and do not trigger target taxonomy admission.

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

Same-tenant iframes, resources outside the captured web boundary, and missing restorable payloads block the current workflow. The default per-dependency capture limit is 10 MiB.

`ClassicWebPartSnapshotReader` only captures common evidence: export XML, ID, title, zone, index, hidden state, and digest. `ClassicWebPartReplayCapabilityPolicy` and `ClassicWebPartActionPlanner` separately assess current Publishing Page replay, List/View dependency closure, and known unsupported types such as RSS Aggregator.

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

The current Enterprise Wiki v1 workflow is intentionally narrow:

- only create-only plans are executable; overwrite/update is refused;
- target pages must be in the root of the target Publishing Pages library;
- the target site collection must already exist; mapped child Webs can be created/recovered, but tenant-level site creation is not implemented;
- child-Web template selection is sealed, but required Feature activation is not inferred or performed; Publishing/Enterprise Wiki, Document ID, Document Set, asset-library, and other Feature prerequisites must already be available or gain explicit plans;
- unique permissions are captured but not restored;
- page-item user, lookup, and taxonomy values still require workflow-specific mappings; List lookup values are mapped through target item receipts, while List taxonomy schema needs a reviewed store/set mapping and the mapped target Terms must already exist;
- Term Group/Set/Term creation and durable many-to-one source taxonomy aliases are not implemented;
- only recognized fields with supported values and compatible target definitions are written;
- version history, Created/Modified/Author/Editor preservation, unique ACLs, workflows/subscriptions/event receivers, and personal Views are not restored;
- custom List View/Web Part `JSLink` or `XslLink` paths block because their exact resource bytes do not yet have a List-rendering-resource materializer;
- supported Views currently replay query, fields, type, row limit, paging, and JSLink; arbitrary full `ListViewXml`, XslLink, hidden state, and every existing-default-view collision shape are not yet exact;
- strict List title collision handling is the default. An explicit target title override is supported; the resumable rename of a specifically selected empty template-created List is not;
- List content-type membership, parent identity, name/description/group/flags, required FieldLinks, and visible order are restored, but arbitrary removal of extra target-runtime FieldLinks and every template-specific CT behavior need live validation;
- dependency materialization happens before page creation; mutation intents and receipts are journaled, but there is no cross-object transaction or automatic global rollback;
- replacements are reviewed and digest-sealed but are case-insensitive text substitutions rather than DOM-aware URL edits;
- supported source-list-bound classic Web Parts are remapped through target Web/List/View receipts; known non-portable types and incomplete bindings remain blocked;
- typed runtime verification requirements are sealed in the plan, but browser automation and its evidence receipt remain outside the library importer;
- a source fence detects capture-time mutation but does not invalidate an export after a later source edit;
- live-tenant behavior still requires environment-specific validation in addition to unit and contract tests.

A future implementation should add an explicit action or a narrower/new workflow instead of silently relaxing one of these blockers.

## Validation expectations

Changes to this family should validate:

- all target frameworks supported by `PnP.Framework` build;
- export, migration package, and receipt JSON round trips;
- snapshot and plan mutation invalidates the corresponding digest;
- exactly one field action exists per captured page field, one dependency action per captured page dependency, one Web Part action per captured Web Part, and one List plan per captured List closure node;
- every non-empty topology, layout-resource/schema, List/site-schema, current item/document/attachment/View, Web Part, and reference ingredient has exactly one non-fallback action;
- dependency releases are unique, belong to `Transform` actions, and name actual required edges;
- topology target analysis covers each mapped Web exactly once, and duplicate Site/Web probes are rejected;
- custom site-content-type closure terminates only at exact runtime IDs and is materialized parent-first;
- lookup cycles and calculated-field dependency cycles block before mutation;
- helper fields use `AddToNoContentType`, explicit List CT order round-trips through target-generated IDs, and read-only runtime values are not replayed;
- import storage success includes final topology and List readback, not only the page file;
- shared `Pages` code has no dependency on `Publishing` or `EnterpriseWiki`;
- publishing lifecycle derivation remains conservative;
- package state agrees with blocker state;
- the Markdown report exposes complete source evidence and every plan action;
- focused classification, replacement, field, lifecycle, Web Part policy, and validation tests pass.
