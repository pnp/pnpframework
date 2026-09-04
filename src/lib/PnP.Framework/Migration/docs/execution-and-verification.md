# Execution and verification

> Status: Draft
> Implementation status: Current execution and receipt model
> Scope: Import admission, mutation journaling, retry, fresh readback, and acceptance

## Execution is transaction-like, not transactional

Each attempted mutation has an intent, an outcome, and a receipt. The overall import has a plan digest, operation ID, ordered steps, terminal status, and verification result. These properties make execution auditable and support safe retry; the importer does not automatically resume from the journal.

They do not create a SharePoint transaction spanning Webs, Lists, schema, files, Web Parts, and a page. There is no atomic commit and no automatic global rollback.

The correct model is:

```text
sealed object action graph
    -> dependency-ordered mutation steps
    -> per-step receipts and runtime IDs
    -> fresh-readback verification
```

## Admission boundary

Import must complete all non-mutating admission checks before its first target write.

### Contract validation

Before an operation ID is created, contract validation checks:

- supported Publishing Page envelope and embedded export schema versions;
- required package/snapshot/plan structure;
- canonical `snapshotDigest` and `planDigest` equality;
- `plan.SourceSnapshotDigest` binding;
- package state and blocker consistency;
- action coverage and identity uniqueness;
- nested topology, schema, List, runtime-verification, and artifact invariants;

Several nested contracts carry their own `schemaVersion`, but current package validation does not uniformly enforce all of those values. A malformed, unsupported, or digest-invalid contract throws before an operation or receipt is created.

### Approval and fresh target admission

After contract validation, admission checks caller-supplied `approvedPlanDigest` equality and performs a read-only target preflight.

Planning probes are review evidence, not execution-time truth. Import freshly inspects critical target state before mutation, including:

- target page and Pages library preconditions;
- layout/resource/schema compatibility;
- topology identity, parentage, template, ownership, and collisions;
- required platform-feature state, activation authority, and promised runtime content types;
- Lists and custom site content types whose owner Web already exists;
- permissions needed for planned materialization.

Objects below a missing but approved Site/Web are reported as a topology prerequisite until that reviewed topology transaction is materialized. This is an ordered action, not a terminal deferred gap. Their frozen source evidence remains valid, and each dependent object must be freshly probed before its own mutation begins.

Topology planning requires a captured direct-parent chain for every executable source Web. It must not invent an intermediate Web from a multi-segment URL. A source Web whose parent chain cannot be read remains non-executable together with only its dependent descendants; unrelated topology branches continue. The workflow may classify that branch as authorization-blocked only when retained wire evidence contains a literal HTTP `401` or `403`. Other incomplete or semantic-access-denied evidence remains a mitigation/RCA result.

Target mapping preserves the complete path relative to the source Site Collection. Run isolation may add the reviewed suffix to the target Site Collection leaf, and a proven collision may add a deterministic suffix only to the colliding Site/Web/List/Page node. Missing nodes use an explicit create action; absence is not a reason to flatten or rename the remaining hierarchy.

### Admission failure

An expected approval or target-admission rejection after successful contract validation returns a receipt with:

- `ExecutionStatus = NotStarted`;
- `MutationStarted = false`;
- a typed `ExecutionAdmissionFailure` code, subject, and message;
- no target mutation.

Admission failure is different from contract validation, which currently throws without a receipt, and from an unexpected exception after execution starts.

Current receipt gap: an admission failure records the sealed package plan digest, not the rejected caller-supplied candidate digest. The typed admission code preserves the reason, but the candidate approval value is not currently auditable from the receipt.

## Operation identity and state

Every invocation that passes contract validation creates a new `OperationId` and is bound to one `PlanDigest`. An invalid contract currently has no operation identity.

`MigrationExecutionStatus` values are:

| Status | Meaning |
| --- | --- |
| `NotStarted` | Admission did not authorize mutation. |
| `Running` | The execution boundary was crossed and work may be partially applied. |
| `Succeeded` | Required execution and library-owned verification completed successfully. |
| `FailedUnexpectedly` | The current implementation uses this for any post-admission attempt that does not end in full mutation-plus-readback success, including an unexpected exception or a required fresh-readback mismatch. Receipts and target state must be inspected before retry. |

Execution status is not the same as source eligibility, plan approval, storage verification, runtime verification, or final acceptance.

The surrounding migration workflow may stop a transaction as authorization-blocked only when it retains a structured wire-level HTTP status code `401` or `403`. A SharePoint `E_ACCESSDENIED` value inside an HTTP 200 CSOM response, a permission-mask prediction, or message text containing “401/403” is evidence for further authorization RCA, not proof of this terminal classification.

## Mutation journal

Before each mutating category, the importer writes `MigrationMutationIntent`:

| Field | Meaning |
| --- | --- |
| `operationId` | Import attempt identity. |
| `planDigest` | Approved plan being executed. |
| `actionId` | Stable step/category identity within the operation. |
| `sequence` | Ordered mutation sequence. |
| `writtenAtUtc` | Time intent was recorded. |
| `description` | Human-readable intended mutation. |

After the operation returns or is proven already satisfied, the importer writes `MigrationMutationReceipt`:

| Field | Meaning |
| --- | --- |
| `operationId`, `planDigest`, `actionId`, `sequence` | Correlate with the intent. |
| `completedAtUtc` | Completion time. |
| `outcome` | `Applied`, `AlreadySatisfied`, or `Failed`. |
| `exchangeIds` | Contract slot for runtime identities made available to later steps. |
| `message` | Outcome details. |

The current recorder leaves `exchangeIds` empty; concrete runtime identity maps are carried by topology and List domain receipts. An intent without a corresponding receipt indicates an interrupted or unobserved completion. The next attempt must freshly inspect the target; it must not assume either success or failure.

## Dependency-ordered execution

The current Publishing Page execution follows the approved dependency graph. At a high level:

1. validate package and fresh admission;
2. materialize/resolve target topology;
3. materialize Page Layout-associated schema, rendering resources, and the layout;
4. materialize approved page-reference dependency artifacts;
5. activate and freshly verify conditional SharePoint platform features before their consuming content-type/List actions;
6. materialize required Lists in lookup dependency order;
7. collect target Web/List/item/View/content-type mappings;
8. create the target Publishing Page;
9. write approved transformed content and supported page fields, then import or rebind classic Web Parts using the runtime mappings;
10. apply the derived target lifecycle;
11. freshly read the supported target closure;
12. produce the import receipt and storage/runtime/acceptance statuses.

The precise step grouping may evolve, but a consumer must not execute an action before its required runtime identity maps are available.

## Ownership and retry

### Owned creation

When a target path is free and the plan authorizes creation, the object is created with deterministic description and migration provenance where supported. Ownership commonly includes:

- a source-qualified original identifier;
- a semantic plan or mapping digest;
- deterministic target location and type.

### Exact reuse

An existing object may be reused only under an explicit domain rule, such as:

- an approved target host;
- a required target-runtime object with compatible identity/schema;
- a migration-owned object with matching original identifier and digest;
- approved stock Page Layout reuse.

Name or title equality alone does not establish ownership.

### Interrupted creation

Topology supports `RecoverInterruptedCreate` when a child Web exists without complete ownership marking but its deterministic path, parent, title, template, and configuration agree with the sealed recovery description. Recovery then applies and verifies ownership provenance.

Other domains must define equally narrow recovery rules before claiming an interrupted unmarked object. The absence of a receipt is not sufficient authority to claim it.

### Planning-time versus actual disposition

A plan may select `CreateOwned` because the target was free. Before execution, another authorized attempt may complete the same exact owned object. Fresh admission may then safely observe `ReuseOwned`.

Receipts record the actual execution-time disposition. Semantic ownership digests exclude mutable observation state where required so this safe create-to-reuse transition does not invalidate the approved intent.

An external topology approval binds the exact plan digest, runner and library binaries, target Site URL, preflight timestamp, and expected action. Execution accepts only safe forward transitions: `CreateMissing` may become interrupted-create recovery, owned reuse, or owned-plan reconciliation; recovery may become owned reuse or reconciliation; reconciliation may become exact owned reuse; exact reuse may remain exact reuse only. A transition to a foreign collision, authorization failure, retryable failure, or a newly required create action is rejected for fresh review before mutation.

## Failure and rollback semantics

If execution fails after one or more successful mutations:

- completed owned objects remain at the target;
- receipts and ownership markers remain recovery evidence;
- the import receipt records `FailedUnexpectedly` and any available completed steps;
- a retry revalidates the same package and freshly probes the target;
- already satisfied exact work may be accepted;
- ambiguous, changed, or unowned state blocks instead of being overwritten.

The importer does not delete completed work merely to recreate a clean starting point. Domain-specific compensation could be designed later, but it must be explicit, independently authorized, and verifiable; it is not implied by the current journal.

## Runtime identity exchange

SharePoint allocates identities that cannot be known at planning time. Domain receipts therefore form runtime catalogs:

- source Web ID -> target Site/Web ID;
- source List ID -> target List ID;
- source item ID -> target item ID;
- source View ID -> target View ID;
- source List-local content type ID -> target List-local content type ID.

Downstream operations consume these maps. Lookup values and List-bound Web Parts must use target-generated IDs. Source IDs remain correlation evidence and must not be forced into target runtime stores.

## Verification levels

Verification has intentionally separate levels:

| Level | Evidence | Authority |
| --- | --- | --- |
| Mutation outcome | CSOM operation returned or state was already satisfied | Execution step only |
| Domain fresh readback | Newly loaded target object graph and bytes | Domain verifier |
| Storage verification | Aggregate required persisted assertions | PnP Framework importer |
| Runtime verification | Browser/runtime behavior for typed requirements | External verifier |
| Acceptance | Required storage and runtime results combined | Receipt/status policy |

The following implications are invalid:

```text
CSOM success => fresh readback passed
page file exists => dependency closure is correct
storage verification passed => browser verification ran
runtime verification pending => storage migration failed
```

## Fresh-readback requirements

Verification must use newly loaded target state after mutation. It should not assert solely against in-memory CSOM objects used for writes.

Current required readback includes, where governed by the plan:

- target Site/Web hierarchy, identity, template/configuration, and ownership;
- List identity, settings, ownership, fields, content types, FieldLinks/order, Views, current values, file bytes, and attachment bytes;
- source-to-target runtime identity mappings;
- classic Web Part count, definition/digest, zone, order, hidden state, and rebound identity properties;
- target page path, file/item/content-type identity, publishing content digest, fields, and lifecycle evidence;
- required Page Layout, resource, and schema state.

Large/binary equality is established using length and SHA-256, with exact byte reads where the verifier owns a byte-stable artifact. An Information Rights Management envelope is not admitted to that verifier: repeated reads can regenerate the DRM transform and encrypted-package streams. Until a reviewed semantic verifier exists, its ingredient remains `Defer`; source comparison uses the retained file identity/version/length plus `cTag` and `QuickXorHash`, while artifact SHA-256 continues to protect the immutable captured envelope itself.

## Verification and acceptance statuses

### Storage

`StorageVerificationStatus`:

- `NotRun`: no valid storage verification result exists;
- `Passed`: all required library-owned assertions passed;
- `Failed`: at least one required storage assertion failed.

### Runtime

`RuntimeVerificationStatus`:

- `NotRun`: no runtime evaluation was performed;
- `NotRequired`: the sealed plan has no required runtime checks;
- `Pending`: runtime checks are required but have not supplied complete passing evidence;
- `Passed`: required runtime results passed;
- `Failed`: at least one required runtime result failed.

The current storage importer emits only `NotRequired` or `Pending` on a successful storage receipt (`NotRun` on failure paths). `Passed` and `Failed` are available to the external runtime receipt model, but are not yet reconciled back into a new import receipt by the library.

### Acceptance

`MigrationAcceptanceStatus`:

- `Pending`: required evidence is incomplete, commonly because runtime verification is pending;
- `Accepted`: all required storage and runtime evidence passed;
- `Rejected`: required storage or runtime evidence failed.

Acceptance is derived from evidence; it must not be used to erase lower-level diagnostics.

## Import receipt as audit boundary

`PublishingPageImportReceipt` records:

- attempt timing and identity;
- admission and mutation-boundary state;
- ordered mutation receipts;
- the approved plan digest;
- actual target page identity;
- actual topology and List runtime maps/dispositions;
- page field, Web Part, content, content type, and lifecycle outcomes;
- aggregate fresh-readback, storage, runtime, and acceptance states;
- warnings and domain diagnostics.

The migration package remains the approved intent. The import receipt is the observed outcome. Neither should be overwritten to make the other appear successful.

## External runtime verification

The plan may seal typed runtime requirements such as page load or Web Part behavior. An external verifier must return results bound to the same plan digest and target identity.

The storage importer must not set runtime verification to `Passed` merely because it emitted a manifest or because the page file exists. Recording a requirement and satisfying it are separate events.

Current gap: `RuntimeVerificationReceipt` is defined, but PnP Framework does not yet validate and merge it with a `PublishingPageImportReceipt` to produce a new final acceptance record. Required runtime work therefore leaves the import receipt at `RuntimeVerificationStatus.Pending` and `MigrationAcceptanceStatus.Pending`; the external workflow must retain both digest-bound records.
