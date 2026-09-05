# Package model

> Status: Draft
> Implementation status: Implemented contracts with explicit extension points
> Applies to: Publishing Page package contracts v2

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
| `PublishingPageExportPackage` | `pnp-publishing-page-export/v2` |
| `PublishingPageMigrationPackage` | `pnp-publishing-page-migration-package/v2` |
| `PublishingPageImportReceipt` | `pnp-publishing-page-import-receipt/v2` |
| `PageArtifactSnapshot` | `pnp-page-artifact/v1` |
| `PageRuntimeSnapshot` | `pnp-page-runtime/v1` |
| `CanonicalPageIngredientGraph` | `pnp-page-ingredient-graph/v1` |
| `RuntimeVerificationManifest` | `pnp-migration-runtime-verification/v1` |
| `RuntimeVerificationReceipt` | `pnp-migration-runtime-verification-receipt/v1` |
| `TaxonomyValueRelationshipSnapshot` | `pnp-taxonomy-value-relationship/v1` |
| `TaxonomyAssetSourceSnapshot` | `pnp-taxonomy-asset-source/v1` |
| `TaxonomyTermGroupMaterializationPlan` | `pnp-taxonomy-termgroup-plan/v1` |
| `TaxonomyTermSetMaterializationPlan` | `pnp-taxonomy-termset-plan/v1` |
| `TaxonomyTermMaterializationPlan` | `pnp-taxonomy-term-plan/v1` |
| `TaxonomyAssetReviewPlan` | `pnp-taxonomy-asset-review-plan/v1` |
| `TaxonomyAssetApprovalManifest` | `pnp-taxonomy-asset-approval/v1` |
| `TaxonomyAssetExecutionAdmission` | `pnp-taxonomy-asset-execution-admission/v1` |
| `TaxonomyAssetMaterializationReceipt` | `pnp-taxonomy-asset-materialization-receipt/v1` |

Nested snapshots and plans have their own schema identifiers where independent evolution or validation is required.

Taxonomy asset evidence is currently emitted as a cohort-level, digest-sealed review sidecar. It is derived from the page bundles but is not yet embedded in `PublishingPageMigrationPackage`. Its separate approval manifest covers every deterministic target TermGroup, TermSet, and Term action. A TermSet action cannot implicitly create its parent TermGroup, and adding a child to an external TermSet requires an additional per-Term flag. Execution produces a journal plus fresh-readback receipt. Page Apply must consume a successful receipt before the two approval boundaries can be joined.

JSON serialization uses camel-case property names, string enum values, explicit nulls, and case-sensitive property names.

## Artifact relationship

```text
PublishingPageExportPackage
├── schemaVersion
├── exportedAtUtc
├── selection
├── selectionDigest
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
├── selection
├── selectionDigest
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
| `selection` | Workflow ID plus versioned validation-cohort assessment. This selects policy; it is not source evidence. |
| `selectionDigest` | SHA-256 over the workflow/cohort selection. Classification edits invalidate the package independently of source evidence. |
| `snapshot` | Complete `PublishingPageCaptureBundle`. |
| `snapshotDigest` | SHA-256 over canonical serialization of the complete snapshot. |

Changing any snapshot evidence after sealing invalidates `snapshotDigest`.

### Capture bundle

`PublishingPageCaptureBundle` contains:

| Field | Captured evidence |
| --- | --- |
| `capturePolicy` | Normalized capture inputs and payload limits. |
| `source` | Page/Web/file/List-item identity, content type, version, size, modified time, and title. |
| `pageArtifact` | Exact source ASPX byte artifact, parsed `Page` directive, availability, and diagnostics. |
| `runtime` | CLR-derived executable adapter, evidence source, resolution state, and diagnostics. |
| `profileSignals` | Non-exclusive Content Type, layout, and field signals. Multiple profile IDs may apply. |
| `ingredientGraph` | Canonical nodes and dependency edges projected over all typed page evidence, including CLR/runtime, layout resources/schema, owner Webs, List/site schema, current items/documents/attachments/Views, Web Parts, and references. |
| `layout` | Publishing Page Layout identity, exact artifact, parsed controls/zones/registrations/resources, and associated schema closure. |
| `publishingPageContent` | Complete source `PublishingPageContent` HTML. |
| `publishingPageContentSha256` | Digest of captured publishing HTML. |
| `fields` | Every returned Pages-library field definition plus typed or raw value evidence. Taxonomy fields additionally carry their exact binding, field-value-set digest, live-resolution state, `TaxonomyHiddenList`/`TaxCatchAll` evidence, and per-value relationship proof. |
| `webParts` | Captured classic Web Part export XML, identity, placement, hidden state, and digest. |
| `listWebPartBindings` | Parsed source Web/List/View bindings and relevant XML/path evidence. |
| `listDependencies` | Required Lists/libraries, settings, fields, site/List content types, Views, current items, folders, files, and attachments. Every returned item field has a value snapshot; unknown runtime types retain best-effort raw evidence and may be marked `Partial`. Binary evidence records whether SharePoint returned an ordinary payload or an IRM envelope; protected documents retain exact artifact bytes plus available `cTag`/`QuickXorHash` logical identity. |
| `listLookupDependencies` | Directed lookup edges used for ordering and cycle detection. |
| `sourceTopology` | Source Site Collection and complete required Web ancestor closure. |
| `dependencies` | Authored references and safe payload evidence. |
| `security` | Permission inheritance and role-assignment evidence. |
| `lifecycle` | Source checkout, file level, moderation, and timestamp evidence. |
| `sourceFence` | Before/after file identity, version, length, and modified-time stability evidence. |
| `blockers` | Capture findings that make the selected workflow non-executable. |
| `warnings` | Review findings that do not independently block planning. |

The bundle is not a list of writes. A value may be captured even when its later plan disposition is evidence-only or blocked.

For an IRM envelope, `artifact.sha256` remains the package-integrity identity of the exact captured response. It is deliberately not interpreted as a stable source-content identity. `logicalContentIdentity.quickXorHash`, together with source file identity/version/length and `contentTag`, supports source-to-source semantic comparison while replay remains `Defer`. The full artifact still participates in `snapshotDigest`, so two faithful captures may have different immutable snapshot digests even when their protected logical document is unchanged.

### Capture-time and plan-time ingredient projections

The snapshot's `ingredientGraph` is immutable capture evidence. Its projection version records the projector semantics used when that snapshot was sealed; validating or planning an older package must never rewrite that graph or change `snapshotDigest`.

Planning independently derives the current canonical graph from the typed snapshot evidence and stores it as `plan.ingredientGraph`. This separates two concerns:

- export validation proves that the captured graph is authentic for its recorded projection semantics;
- plan validation proves that the actions and dependency closure match the current projector and policy.

Legacy export packages whose graph has no `projectionVersion` are validated against the legacy projector. A current plan over that evidence stores the current versioned projection in `plan.ingredientGraph`, while the embedded snapshot remains byte-for-byte and digest-equivalent to the export. Import validates both boundaries. This permits projector evolution without either silently accepting a tampered old graph or invalidating authentic frozen evidence.

## Target-specific migration package

`PublishingPageMigrationPackage` embeds the source snapshot and adds the complete reviewed target intent.

| Field | Meaning |
| --- | --- |
| `schemaVersion` | Migration-package envelope version. |
| `plannedAtUtc` | Time target analysis and plan sealing completed. |
| `exportSchemaVersion` / `exportedAtUtc` | Provenance of the embedded export. |
| `state` | `ApprovalReady` when executable; `MitigationPending` when more evidence/capability work is queued; `AuthorizationBlocked` only for literal HTTP 401/403; `Invalid` for an inconsistent action graph. `Draft` is available during construction and legacy `Blocked` is no longer emitted. |
| `selection` | Exact workflow/cohort selection copied from the export and validated again at planning/import. |
| `selectionDigest` | Must continue to match the embedded selection and its policy-derived assessment. |
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
| `taxonomyRelationshipActions` | Exactly one action for every captured taxonomy value. A selected field seals the exact target field/text-field binding and reuses an exact live-in-bound Term, reproduces an exact live-outside-bound or dangling relationship, or defers for mitigation. An unselected field uses `RetainEvidenceOnly`, which preserves sealed source evidence without making a target claim. A relationship action never authorizes Term creation, repair, or substitution; any exact missing source asset must be prepared through a separately reviewed taxonomy asset plan before page admission. |
| `dependencyActions` | Exactly one result for every captured governed reference. |
| `topology` | Source Site/Web to target Site/Web mapping and topology semantic digest. |
| `topologyTargetAnalysis` | Target existence, identity, parent, template, ownership, disposition, and issues for each mapped Site/Web. |
| `listMigration` | Ordered per-List plans, conditional platform-feature requirements, field/View/site-content-type actions, target probes, issues, and digests. Each feature requirement seals its ID, scope, dependency order, consuming and promised content-type IDs, and target Site Collection. |
| `webPartActions` | Copy, rebind-after-materialization, or defer-for-mitigation result for each captured Web Part. |
| `replacements` | Approved source-to-target text substitutions. |
| `expectedPublishingPageContentSha256` | Expected post-replacement publishing-content digest. |
| `storageAssertions` | Required storage-level expectations. |
| `runtimeVerification` | Typed requirements for an external verifier. Presence does not imply execution. |
| `ingredientGraph` | Current versioned projection derived from the immutable typed snapshot evidence for this plan. It may differ from the capture-time graph only through an explicit projector-version change; planning does not rewrite the snapshot. |
| `ingredientActions` | Exactly one semantic capability/disposition, target, policy, dependency-release list, and verification list for every non-empty canonical ingredient. Releases are valid only on a `Transform` and must name a real required dependency edge. |
| `migrationOutcome` | Evaluated aggregate: `Exact`, `ExecutableWithTransform`, `ExecutableWithLoss`, `MitigationPending`, `AuthorizationBlocked`, `Invalid`, or `Unknown`. Legacy `Blocked` is no longer emitted. |
| `ingredientIssues` | Recomputed dependency-closure and action-coverage issues. |
| `blockers` / `warnings` | Plan-wide findings. `IsExecutable` requires both an empty blocker list and an executable ingredient outcome. |

The plan contains nested actions rather than a flat transaction list. Dependency ordering and runtime identity exchange determine execution order.

Package-level `MitigationPending` means that this sealed target transaction is not safe to execute yet, but it is a nonterminal work item: the migration loop continues with evidence collection, RCA, capability implementation, re-capture, replanning, and verification. `AuthorizationBlocked` is the only authorization stop and requires retained literal wire-level HTTP `401` or `403`. `Invalid` identifies an inconsistent proposed action graph and also returns to RCA rather than being treated as authorization. The legacy ambiguous `Blocked` state is retained only for contract compatibility and is never emitted by new planning.

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
| `fieldResults` | Per-page-field execution result, including target-local taxonomy materialization receipts where applicable. |
| `taxonomyRelationshipsMatched` / `taxonomyRelationshipResults` | Aggregate and per-executed-value fresh readback of field binding, page value, live/absent Term state, hidden-list identity, and `TaxCatchAll`. Evidence-only relationships are not target assertions. |
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
- workflow ID and validation-cohort assessment;
- canonical selection-digest equality and workflow-policy re-evaluation from source evidence;
- source ASPX artifact identity plus byte length/digest;
- runtime, profile-signal, and canonical ingredient-graph structure plus deterministic re-derivation from the typed source evidence;
- canonical snapshot and plan digest equality;
- plan-to-snapshot digest binding;
- package state consistent with blockers;
- exactly one governed action per captured page field, dependency, and Web Part;
- exactly one List plan per captured List dependency;
- complete and unique topology target-analysis coverage;
- List identities, admitted probes for executable Lists, content-type closure shape, and per-List/plan-set digests;
- unique runtime-verification requirement IDs;
- unique ingredient/action IDs, deterministic action re-projection from typed domain plans, complete non-empty ingredient action coverage, valid graph edges, validated transform-only dependency releases, recomputed dependency closure, and aggregate outcome equality;
- artifact availability and digest validity when external artifacts are required.

Current validation gaps include uniform checking of every nested `schemaVersion` value and independent semantic re-derivation of all sealed List ordering and target-planning decisions. Missing nested List/schema/View decisions already project to blocking ingredient actions; those remaining gaps do not authorize Import to repair or reinterpret the package.

A contract-validation failure throws before target admission and currently produces no import receipt. Approval and target-admission failures after successful contract validation return typed zero-mutation receipts.

## Human-readable report

The Markdown migration report is a bounded projection of `PublishingPageMigrationPackage`. It is designed for plan review and includes source evidence, CLR runtime resolution, profile/cohort classification, all ingredient nodes and edges, every ingredient and typed-domain action, target probes, issues, expected assertions, and approval digests.

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
