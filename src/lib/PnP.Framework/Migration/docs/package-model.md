# Package model

> Status: Draft
> Implementation status: Implemented contracts with explicit extension points
> Applies to: Publishing Page package contracts v1

## Purpose

The package model separates three questions:

1. What was captured from the source?
2. What is approved for one target?
3. What actually happened and was verified?

No single artifact answers all three questions. The complete migration record is a chain of digest-sealed intent and observed outcome:

```text
Export package
    -> Migration package
        -> Import receipt
            -> optional runtime verification receipt
```

The migration package embeds the source snapshot, so Import does not need to reconnect to the source. The import receipt records actual target state; it does not rewrite the approved package.

## Versioned artifacts

| CLR contract | Schema |
| --- | --- |
| `PublishingPageExportPackage` | `pnp-publishing-page-export/v1` |
| `PublishingPageMigrationPackage` | `pnp-publishing-page-migration-package/v1` |
| `PublishingPageImportReceipt` | `pnp-publishing-page-import-receipt/v1` |
| `RuntimeVerificationManifest` | `pnp-migration-runtime-verification/v1` |
| `RuntimeVerificationReceipt` | `pnp-migration-runtime-verification-receipt/v1` |

Nested snapshots and plans have their own schema identifiers where independent evolution or validation is required.

JSON serialization uses camel-case property names, string enum values, explicit nulls, and case-sensitive property names.

## Artifact relationship

```text
PublishingPageExportPackage
├── schemaVersion
├── exportedAtUtc
├── snapshot
└── snapshotDigest
          |
          | exact snapshot and export provenance are embedded
          v
PublishingPageMigrationPackage
├── schemaVersion
├── plannedAtUtc
├── exportSchemaVersion
├── exportedAtUtc
├── state
├── snapshot
├── plan
├── snapshotDigest
├── planDigest
└── report
          |
          | caller presents approvedPlanDigest
          v
PublishingPageImportReceipt
├── operation and admission outcome
├── approvedPlanDigest
├── mutation receipts
├── actual target identities
├── domain materialization receipts
├── fresh storage verification
└── runtime/acceptance status
```

## Source export package

`PublishingPageExportPackage` is the portable result of source-only capture.

| Field | Meaning |
| --- | --- |
| `schemaVersion` | Export envelope version. |
| `exportedAtUtc` | Time capture and package sealing completed. |
| `snapshot` | Complete `PublishingPageCaptureBundle`. |
| `snapshotDigest` | SHA-256 over canonical serialization of the complete snapshot. |

Changing any snapshot evidence after sealing invalidates `snapshotDigest`.

### Capture bundle

`PublishingPageCaptureBundle` contains:

| Field | Captured evidence |
| --- | --- |
| `sourceProfile` | Profile that classified the source, currently Enterprise Wiki. |
| `capturePolicy` | Normalized capture inputs and payload limits. |
| `source` | Page/Web/file/List-item identity, content type, version, size, modified time, and title. |
| `layout` | Publishing Page Layout identity, exact artifact, parsed controls/zones/registrations/resources, and associated schema closure. |
| `publishingPageContent` | Complete source `PublishingPageContent` HTML. |
| `publishingPageContentSha256` | Digest of captured publishing HTML. |
| `fields` | Every returned Pages-library field definition plus typed or raw value evidence. |
| `webParts` | Captured classic Web Part export XML, identity, placement, hidden state, and digest. |
| `listWebPartBindings` | Parsed source Web/List/View bindings and relevant XML/path evidence. |
| `listDependencies` | Required Lists/libraries, settings, fields, site/List content types, Views, current items, folders, files, and attachments. Every returned item field has a value snapshot; unknown runtime types retain best-effort raw evidence and may be marked `Partial`. |
| `listLookupDependencies` | Directed lookup edges used for ordering and cycle detection. |
| `sourceTopology` | Source Site Collection and complete required Web ancestor closure. |
| `dependencies` | Authored references and safe payload evidence. |
| `security` | Permission inheritance and role-assignment evidence. |
| `lifecycle` | Source checkout, file level, moderation, and timestamp evidence. |
| `sourceFence` | Before/after file identity, version, length, and modified-time stability evidence. |
| `blockers` | Capture findings that make the current exact profile non-executable. |
| `warnings` | Review findings that do not independently block planning. |

The bundle is not a list of writes. A value may be captured even when its later plan disposition is evidence-only or blocked.

## Target-specific migration package

`PublishingPageMigrationPackage` embeds the source snapshot and adds the complete reviewed target intent.

| Field | Meaning |
| --- | --- |
| `schemaVersion` | Migration-package envelope version. |
| `plannedAtUtc` | Time target analysis and plan sealing completed. |
| `exportSchemaVersion` / `exportedAtUtc` | Provenance of the embedded export. |
| `state` | `ApprovalReady` when the plan has no blockers, otherwise `Blocked`; `Draft` is available while constructing a package. |
| `snapshot` | Exact source evidence used to make the plan. |
| `plan` | Target mappings, actions, probes, expected assertions, and issues. |
| `snapshotDigest` | Must still match the embedded snapshot. |
| `planDigest` | SHA-256 over canonical serialization of the complete plan, including policy, target mappings, planning probes, actions, issues, and expected assertions. |
| `report` | Report metadata. The Markdown report is generated from the package. |

`planDigest` is the review and approval token. Import requires the caller to present the exact approved value.

### Migration plan

`PublishingPageMigrationPlan` is the root target-specific action graph.

| Field | Meaning |
| --- | --- |
| `sourceSnapshotDigest` | Binds the plan to one exact source snapshot. |
| `sourceWebUrl` / `sourcePageServerRelativeUrl` | Source boundary used by reviewed mappings. |
| `targetWebUrl` / `targetWebServerRelativeUrl` / `targetPageServerRelativeUrl` | Approved target location. |
| `pageLayoutName` | Approved target layout selected by the layout plan. |
| `operation` | Currently executable as `CreatePage`; deferred-field recovery remains represented but is not executable. |
| `targetLifecycle` / `lifecycleReason` | Derived Draft/Published result and rationale. |
| `createOnly` | Current policy requires a new target page and blocks an existing target page. |
| `planningPolicy` | Normalized planning inputs included in the approval boundary. |
| `targetProbe` | Planning-time target Web, Pages library, lifecycle, content-type, layout, page, and dependency observations. |
| `layoutMaterialization` | Stock-reuse or deterministic owned-layout action and required resources/schema. |
| `layoutTargetProbe` | Detailed target evidence for layout bytes, schema, permissions, and resources. |
| `layoutAdmission` | Typed eligibility result and issues for the layout closure. |
| `fieldActions` | Exactly one result for every captured governed page field. |
| `dependencyActions` | Exactly one result for every captured governed reference. |
| `topology` | Source Site/Web to target Site/Web mapping and topology semantic digest. |
| `topologyTargetAnalysis` | Target existence, identity, parent, template, ownership, disposition, and issues for each mapped Site/Web. |
| `listMigration` | Ordered per-List plans, field/View/site-content-type actions, target probes, issues, and digests. |
| `webPartActions` | Copy, rebind-after-materialization, or block result for each captured Web Part. |
| `replacements` | Approved source-to-target text substitutions. |
| `expectedPublishingPageContentSha256` | Expected post-replacement publishing-content digest. |
| `storageAssertions` | Required storage-level expectations. |
| `runtimeVerification` | Typed requirements for an external verifier. Presence does not imply execution. |
| `blockers` / `warnings` | Plan-wide findings. `IsExecutable` is derived from an empty blocker list. |

The plan contains nested actions rather than a flat transaction list. Dependency ordering and runtime identity exchange determine execution order.

## Target information

Target-related fields have different authority and lifetime:

| Target concept | Stored in | Question answered |
| --- | --- | --- |
| Mapping/specification | Plan | Where should this source object go? |
| Planning probe | Plan | What existed when the plan was created? |
| Fresh admission probe | Recomputed at Import | Is the target still safe for this approved plan? |
| Materialization receipt | Import receipt | What target object and runtime IDs were actually produced or reused? |
| Verification result | Import receipt | Does fresh persisted state satisfy the approved expectation? |

The package stores planning probes because they are review evidence. Import must not rely on them as current truth; critical facts are freshly inspected before mutation. Fresh admission keeps those new observations transient and does not rewrite the sealed plan.

## Import receipt

`PublishingPageImportReceipt` is an observed outcome for one import attempt.

| Field group | Meaning |
| --- | --- |
| `startedAtUtc`, `completedAtUtc`, `operationId` | Attempt identity and interval. |
| `executionStatus` | `NotStarted`, `Running`, `Succeeded`, or `FailedUnexpectedly`. |
| `admissionFailure`, `mutationStarted` | Explains a zero-mutation rejection and whether execution crossed the mutation boundary. |
| `steps` | Ordered `MigrationMutationReceipt` records. |
| `approvedPlanDigest` | Approval token used by the admitted execution. On the successful path this is the caller-supplied value. |
| Target page identity fields | Freshly observed Web URL, page path, file ID, item ID, content type, and version. |
| Lifecycle fields | Planned lifecycle and actual file/check-out/moderation evidence. |
| Content fields | Expected and persisted publishing-content digests. |
| Web Part fields | Imported count and per-part fresh-readback results. |
| `topologyMaterialization`, `topologyMatched` | Runtime Web mappings, actual dispositions, mapping digests, diagnostics, and topology readback. |
| `listMaterializations`, `listsMatched` | Runtime Web/List/item/View/content-type maps, actual List dispositions, verified counts, diagnostics, and closure readback. |
| `fieldResults` | Per-page-field execution result. |
| `freshReadbackPassed` | Aggregate required readback result. |
| Storage/runtime/acceptance statuses | Distinguish library-owned verification from external runtime work and final acceptance. |
| `warnings` | Non-fatal execution or verification findings. |

The receipt describes one attempt. A later retry has a new `operationId` and may observe `AlreadySatisfied`, `ReuseOwned`, or recovery dispositions for work created by an earlier attempt.

Current receipt gaps:

- malformed, unsupported, or digest-invalid packages throw before an operation ID or receipt is created;
- after contract validation, the zero-mutation admission-failure factory records the package's sealed `planDigest` in `approvedPlanDigest`, because it is not passed the rejected caller value. `admissionFailure.code` still distinguishes `PlanDigestNotApproved`, but the rejected candidate digest is not retained. A future receipt revision should either record both values or leave the approved value unset when approval failed.

## Runtime verification receipt

The plan's `RuntimeVerificationManifest` contains typed requirements. An external runner returns `RuntimeVerificationReceipt` with:

- the same approved `planDigest`;
- a target identity;
- completion time;
- one result per requirement;
- aggregate runtime status.

Runtime verification is separate because the storage importer does not own a browser session or claim visual/interactive acceptance.

The receipt contract exists, but the current library does not yet expose a reconciler that validates this external receipt against a `PublishingPageImportReceipt` and emits a new final acceptance record. When required runtime checks exist, the storage import receipt therefore remains `Pending`; external orchestration must currently preserve and correlate both receipts.

## Digest boundaries

```text
canonical(snapshot)
    -> snapshotDigest

canonical(plan bound to snapshotDigest)
    -> planDigest

caller-supplied approvedPlanDigest
    == package.planDigest
    -> import may proceed after fresh admission
```

The snapshot digest and plan digest serve different review questions:

- `snapshotDigest`: Did the captured evidence change?
- `planDigest`: Did any content inside the sealed plan change?

The boundaries are deliberately narrower than the envelopes:

- `snapshotDigest` covers `snapshot`; it does not cover export `schemaVersion` or `exportedAtUtc`;
- `planDigest` covers `plan`; it does not cover migration-package timestamps, envelope versions, `state`, or `report` metadata;
- `plan.sourceSnapshotDigest` binds the plan to the exact embedded snapshot digest;
- the package validator derives `state` from plan blockers, while the human-readable report remains a non-authoritative projection that can be regenerated from the package.

Some nested domains also carry semantic ownership digests. For example, a List's semantic digest normalizes execution-time target observation state so an approved `CreateOwned` intent can later be recognized as the same owned object and recorded as `ReuseOwned`. This does not mutate or weaken the top-level `planDigest`, which still covers the complete planning-time plan.

Execution-time runtime IDs do not belong in `planDigest` because SharePoint may allocate them only after mutation. They belong in receipts and are correlated by source identity and plan digest.

## Package validation invariants

Before target mutation, current validators require at least the following:

- supported Publishing Page envelope and embedded export schema versions;
- non-null required aggregates;
- canonical snapshot and plan digest equality;
- plan-to-snapshot digest binding;
- package state consistent with blockers;
- exactly one governed action per captured page field, dependency, and Web Part;
- exactly one List plan per captured List dependency;
- complete and unique topology target-analysis coverage;
- List identities, admitted probes for executable Lists, content-type closure shape, and per-List/plan-set digests;
- unique runtime-verification requirement IDs;
- artifact availability and digest validity when external artifacts are required.

Current validation gaps include uniform checking of nested `schemaVersion` values and independent re-derivation of the sealed List order and all nested List field/View action coverage. Those gaps do not authorize Import to repair or reinterpret the package.

A contract-validation failure throws before target admission and currently produces no import receipt. Approval and target-admission failures after successful contract validation return typed zero-mutation receipts.

## Human-readable report

The Markdown migration report is a bounded projection of `PublishingPageMigrationPackage`. It is designed for plan review and includes source evidence, actions, target probes, issues, expected assertions, and approval digests.

It is not the authoritative package and it is not a post-import verification report. Complete large values remain in JSON or the content-addressed artifact store. Actual execution and verification belong to `PublishingPageImportReceipt`.

## Example trace

For a source List-bound Web Part, the artifact chain is conceptually:

```text
snapshot.listWebPartBindings
    source Web/List/View IDs

snapshot.listDependencies
    complete required List and View evidence

plan.topology + plan.listMigration
    target Web path, target List path, creation/reuse actions, probes

plan.webPartActions
    RebindListAfterMaterialization

receipt.listMaterializations
    actual target Web/List/View IDs

receipt.webPartResults
    rebound export, placement, and fresh-readback result
```

The source List/View IDs remain evidence and correlation keys. The target IDs are execution outputs; downstream XML rewriting uses those target values instead of attempting to preserve source runtime identities.
