# Purpose and scope

> Status: Draft
> Implementation status: Current design with explicit gaps
> Scope: `PnP.Framework.Migration`

## Problem statement

A SharePoint page is not an isolated file. Its usable state can depend on its owning Site and Web, publishing Page Layout, site fields and content types, Lists and Views, List items and files, classic Web Parts, authored resources, taxonomy identities, lifecycle settings, and runtime-generated identifiers.

A direct source-to-target copy can therefore report success while leaving a page that cannot render, points back to the source, references the wrong List or View, or has silently lost evidence that the current implementation did not understand.

The migration subsystem treats cross-site movement as a staged and reviewable reproduction of an object graph:

```text
capture source evidence
    -> inspect one specific target
    -> choose explicit actions
    -> approve a sealed plan
    -> execute in dependency order
    -> verify persisted target state
```

The initial profile is Enterprise Wiki Publishing Pages. The architecture is intentionally broader: Topology, Schema, Lists, shared Page capabilities, execution, packaging, and verification are reusable migration domains rather than Enterprise Wiki implementation details.

## Goals

### Preserve source evidence

Capture should retain what SharePoint returned, including typed values, raw fallback representations, diagnostics, and unavailable states. Evidence that cannot be restored today should remain available for a future mapper or importer.

The design follows this rule:

> Capture complete evidence; restore only understood and approved behavior.

### Separate source capture from target policy

Export must be source-only. It must not require a target connection, select target paths, or decide whether a target object should be created or reused.

Planning is target-specific and read-only. The same export may be planned independently for different targets without recapturing the source.

### Make target decisions reviewable

Every captured object governed by the current profile must receive an explicit plan result. Depending on the domain, that result may create, reuse, recover, transform, apply, preserve as evidence, skip conservatively, delegate, or block.

The plan must retain enough target evidence and reasoning for a reviewer to understand why the action was selected.

### Bind approval to exact content

The source snapshot and target plan have independent canonical digests:

- `snapshotDigest` identifies the exact captured source evidence;
- `planDigest` identifies the exact target-specific decisions and assertions.

Import accepts an approved plan digest. Editing a target path, action, mapping, lifecycle result, blocker, or assertion invalidates that approval boundary.

### Preserve object ownership and identity boundaries

The importer must distinguish approved hosts, migration-owned objects, runtime-provided objects, and unrelated target objects. A matching name is not sufficient proof of identity.

Created or reused objects use source-qualified identifiers and semantic digests where the domain supports migration ownership. Planning never overwrites an ambiguous or unowned collision: it keeps the complete mapped relative path and allocates a stable suffix only at the colliding Web, List/library, or Page leaf. If the sealed final path becomes occupied after approval, Apply rejects the stale plan and the workflow replans instead of silently retargeting.

### Support safe retry

SharePoint does not provide a transaction spanning Webs, Lists, files, and pages. The subsystem therefore uses deterministic target descriptions, target re-probing, ownership markers, mutation intents, receipts, and fresh readback to make completed work recognizable on a later attempt.

Safe retry means resuming or accepting already-satisfied owned work. It does not mean globally rolling back a partially completed migration.

### Verify persisted behavior

A successful CSOM request is not sufficient proof of migration success. Verification must freshly read the target and compare persisted identity, schema, settings, values, bytes, bindings, and lifecycle against the approved plan.

Storage verification and browser/runtime acceptance are separate. The library owns storage readback; an external browser-capable runner may satisfy typed runtime requirements.

### Reuse object-domain capabilities

Future Wiki Page and Web Part Page profiles should compose the same Topology, Schema, List, field, reference, and classic Web Part capabilities where their semantics match. A profile should coordinate shared domains without becoming their owner.

## Non-goals

### Full SharePoint backup and restore

The subsystem is not a farm, tenant, Site Collection, or compliance backup product. It does not promise complete recovery of every SharePoint property or service-side behavior.

### A single universal page contract

Publishing Pages, Wiki Pages, and Web Part Pages have different storage models. Shared capabilities belong under `Migration.Pages`, but each page family should own its real aggregate, package, lifecycle interpretation, and verifier.

### Global atomicity or automatic rollback

There is no cross-object commit boundary covering Site/Web creation, schema, Lists, files, Web Parts, and a page. If a later step fails, successfully completed earlier objects are not automatically deleted.

### Preserving source runtime identifiers

Target SharePoint allocates Web, List, View, item, List-local content type, and taxonomy cache identifiers. The importer records target identities and rewrites downstream consumers; it does not force source runtime IDs into the target.

### Guessing unsupported mappings

Missing taxonomy, user, security, resource, or object mappings do not authorize best-effort writes. When correctness depends on a mapping, the plan must contain a reviewed mapping or remain blocked/evidence-only.

### Browser automation inside the storage importer

The storage importer may emit a runtime verification manifest, but it does not claim to have loaded or visually validated the page. Browser automation and its evidence remain an external capability.

### Tenant-wide provisioning in the current implementation

The current topology executor requires an existing target Site Collection. It can create or recover approved child Webs, but tenant-scoped Site Collection provisioning remains an explicit gap.

## Current product boundary

The current implementation focuses on cross-site reproduction of classic Enterprise Wiki Publishing Pages and their required closure:

- source Site Collection and Web ancestry;
- Publishing Page identity, content, lifecycle, fields, permissions evidence, and source stability fence;
- stock or custom Page Layout evidence and required schema/resources;
- classic Web Parts and parsed List/View bindings;
- required Lists, lookup dependencies, fields, content types, Views, current items, folders, files, and attachments;
- explicit target mappings, target collision probes, ownership rules, and blockers;
- dependency-ordered target materialization;
- source-to-target runtime identity catalogs;
- final topology, List, Web Part, field, page-content, content-type, and lifecycle readback;
- optional external runtime verification requirements.

## Explicit gaps

The design intentionally exposes rather than hides the following gaps:

- target Site Collection creation and tenant-scoped admission;
- required Feature activation and template-specific provisioning prerequisites;
- uniform enforcement of every nested contract's `schemaVersion` during package validation;
- a typed operation/receipt for malformed or digest-invalid packages, which currently fail before an import operation is created;
- retention of the caller-supplied approval candidate when plan-digest admission fails;
- guaranteed lossless serialization of every unknown CSOM List-item value; current capture retains each returned field as typed or best-effort raw evidence and marks unsupported representations `Partial`;
- resumable claiming or renaming of selected template-created Lists;
- materialization of custom List View/Web Part `JSLink` and `XslLink` resources;
- Term Group, Term Set, and Term provisioning plus source-alias retention;
- full fidelity for every `ListViewXml` and template-specific View behavior;
- exact removal of arbitrary extra target List content type FieldLinks;
- version history, source audit identity/timestamps, unique ACL replay, workflows, subscriptions, and event receivers;
- personal View recreation;
- library-owned validation and reconciliation of an external `RuntimeVerificationReceipt` into a final combined acceptance record;
- browser DOM, interaction, and visual acceptance.

An explicit gap must not be removed by adding a permissive write alone. It should be closed by adding evidence, an action, target admission, execution ownership, a receipt, and verification together.

## Definition of success

A migration attempt is successful at the storage layer only when:

1. the supplied snapshot and plan digests validate;
2. fresh admission confirms that the target remains compatible with the approved plan;
3. all required mutations complete or are proven already satisfied;
4. runtime-generated target identities are recorded for downstream consumers;
5. fresh readback satisfies every required storage assertion;
6. the import receipt records `StorageVerificationStatus.Passed`.

Final acceptance may remain pending when the sealed plan contains required runtime verification that has not yet been executed.
