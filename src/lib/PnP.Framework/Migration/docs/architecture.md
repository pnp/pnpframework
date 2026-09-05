# Architecture

> Status: Draft
> Implementation status: Current design with explicit gaps
> Scope: `PnP.Framework.Migration`

## Architectural shape

Migration is a staged compiler-and-executor pipeline rather than a direct copy operation:

```text
Source SharePoint
      |
      | Capture: source read only
      v
Source snapshot bundle
      |
      | Package and seal
      v
Export package
      |
      | Plan: target read only
      v
Target-specific object action graph
      |
      | Seal and review
      v
Migration package
      |
      | Fresh admission and execute
      v
Mutation journal + import receipt
      |
      | Fresh storage readback
      v
Storage verification
      |
      | Optional external runner
      v
Runtime verification and final acceptance
```

The snapshot is the input program, the planner compiles it for one target, and the importer executes only the sealed result. The analogy is limited: SharePoint mutations are not an atomic transaction and the importer does not provide automatic global rollback.

## Stage boundaries

| Stage | Source connection | Target connection | Target writes | Primary output |
| --- | --- | --- | --- | --- |
| Capture/export | Required | None | None | `PublishingPageExportPackage` |
| Planning | Not required; uses snapshot | Required, read-only | None | `PublishingPageMigrationPackage` and review report |
| Package validation | None | Supplied but not queried | None | Validated contract, or an exception before an operation/receipt exists |
| Import admission | None | Required, read-only | None | Admission result or authorization to execute |
| Execution | None | Required, read/write | Yes | Mutation receipts and runtime identity maps |
| Storage verification | None | Required, fresh read | None | Verified `PublishingPageImportReceipt` |
| Runtime verification | None | External browser/runtime access | Outside library | `RuntimeVerificationReceipt` |

No stage may silently borrow responsibilities from an earlier stage. In particular:

- capture must not choose a target;
- planning must not mutate a target;
- import must not invent decisions missing from the plan;
- verification must not treat a mutation response as persisted evidence.

## Artifact flow

The Publishing Page family currently defines three primary versioned artifacts:

```text
PublishingPageExportPackage
    selection + snapshot + snapshotDigest
             |
             | embedded without semantic mutation
             v
PublishingPageMigrationPackage
    selection + snapshot + plan + snapshotDigest + planDigest
             |
             | approvedPlanDigest
             v
PublishingPageImportReceipt
    execution + actual target identities + verification
```

The plan may also contain a `RuntimeVerificationManifest`. An external verifier returns a separate digest-bound `RuntimeVerificationReceipt`; runtime results do not retroactively change the migration package.

See [Package model](package-model.md) for the field-level artifact structure.

## Object graph and ordering

A page migration is an object graph. The current cross-site dependency direction is:

```text
source Site/Web ancestor closure
    -> approved target Site/Web map
    -> Page Layout associated schema, rendering resources, and layout
    -> approved page-reference dependency artifacts
    -> lookup-provider Lists before lookup-consuming Lists
         -> required SharePoint platform features and their promised runtime content types
         -> required site content type closure
         -> List identity, settings, fields, and List-local content types
         -> folders, items, documents, attachments, and Views
    -> target Web/List/View/item/content-type identity catalogs
    -> create the Publishing Page
    -> write approved content and fields, then import/rebind classic Web Parts
    -> apply the derived Publishing lifecycle
    -> final topology/List/page fresh readback
```

This ordering is data, not an implementation accident. For example, a List-bound Web Part cannot be correctly rewritten until target Web, List, and View IDs are known. A lookup item value cannot be written until the dependency List has produced a source-item-to-target-item map.

## Domain ownership

Folders and namespaces represent object domains or public workflow boundaries. They are not broad technical-role buckets.

| Namespace | Owns |
| --- | --- |
| `Migration.Diagnostics` | Typed issues and severity. |
| `Migration.Evidence` | Evidence availability, lineage, and recovery fidelity. |
| `Migration.Execution` | Operation state, mutation intents, mutation receipts, and journals. |
| `Migration.Features` | Conditional SharePoint platform-feature requirements, target probes, activation, and runtime-contract verification. |
| `Migration.Packaging` | Artifact references, canonical digests, and artifact stores. |
| `Migration.Topology` | Source Site/Web evidence, target mapping, ownership, Web materialization, and runtime Web IDs. |
| `Migration.Schema.Fields` | Portable site-field evidence, planning, target probes, and schema materialization. |
| `Migration.Schema.ContentTypes` | Site content type closure, planning, target probes, materialization, and verification. |
| `Migration.Lists` | Required List closure, lookup ordering, List-local schema, current content, Views, receipts, and verification. |
| `Migration.Taxonomy` | Taxonomy relationship evidence, reviewed mappings, required TermSet/Term asset closure, provenance, target classification, and asset planning. |
| `Migration.Pages` | Page-wide identity, evidence, fields, references, security, classic Web Parts, and shared mechanics. |
| `Migration.Pages.Markup` | Exact ASPX artifacts and parsed Page-directive evidence. |
| `Migration.Pages.Runtime` | CLR-first runtime-adapter resolution. |
| `Migration.Pages.Profiles` / `Pages.Cohorts` | Non-exclusive profile signals and versioned validation-cohort results. |
| `Migration.Pages.Ingredients` | Canonical ingredient nodes, dependency edges, semantic actions, and aggregate outcome evaluation. |
| `Migration.Pages.ClassicWebParts.Planning` | Current classic Web Part replay capability and action planning. |
| `Migration.Pages.Publishing` | Publishing aggregate, Page Layout, target lifecycle, workflow policy, package, report, execution, and verification. |
| `Migration.Pages.Publishing.EnterpriseWiki` | Thin Enterprise Wiki v1 API facade, discovery, and workflow-specific file naming. |
| `Migration.Verification` | Storage/runtime status and external runtime verification contracts. |

Dependencies point inward:

```text
Enterprise Wiki v1 facade and workflow policy
    -> Publishing Page family
        -> shared Page capabilities
            -> Topology / Schema / Lists / Execution / Packaging / Verification
```

Shared layers must not reference a narrower family or profile. A future Wiki Page family should compose shared page and object domains; it should not inherit a Publishing Page snapshot merely to reuse field or Web Part logic.

## Capture architecture

Capture produces evidence, not actions. Each reader should record:

- stable source identity where SharePoint exposes one;
- the returned value or schema;
- availability/fidelity state;
- diagnostics when the value is missing, denied, ambiguous, unsupported, or failed;
- raw fallback evidence when typed serialization is not available;
- content digests and artifact references for large or binary values;
- a source stability fence when concurrent source mutation would invalidate a coherent snapshot.

Capture aggregates these object snapshots into `PublishingPageCaptureBundle`, retains the exact source ASPX artifact, resolves the CLR runtime, emits non-exclusive profile signals, and projects a canonical ingredient dependency graph. Core, Layout, Topology, List-schema, List-content, Web Part, and Reference projectors own their object-specific nodes and edges; the top-level Publishing projector only composes them. Unknown evidence remains in the bundle even when the current planner later chooses `Delegate`, `EvidenceOnly`, a conservative skip, or a blocker.

## Planning architecture

Planning has two inputs:

1. a sealed source snapshot;
2. read-only observations and policy for one target.

The planner produces typed domain plans plus one canonical `PageIngredientAction` for every non-empty ingredient. Matching object-owned action projectors derive canonical actions from those typed plans; package validation re-projects them instead of trusting a caller-edited summary. There is no universal CLR base class for all executable operations because Web creation, field application, reference rewriting, and List content materialization have different domain semantics. They share conceptual requirements:

- source identity;
- target locator or mapping;
- action/disposition;
- target probe and preconditions;
- dependency relationship;
- reason, diagnostics, or typed issues;
- expected persisted state;
- semantic digest where ownership or approval requires it.

Canonical ingredient dispositions are `Preserve`, `Transform`, `Substitute`, `Drop`, `Delegate`, `Defer`, and `Block`. `Defer` means nonterminal mitigation work: the ingredient is not executable in the current transaction, but it remains in the RCA/re-capture/re-plan queue. `Block` is reserved exclusively for an ingredient bound to retained, digest-valid literal wire HTTP `401` or `403` evidence. A retained consumer cannot lose a required dependency unless its `Transform` explicitly releases a real required edge. The evaluator derives `Exact`, `ExecutableWithTransform`, `ExecutableWithLoss`, `MitigationPending`, `AuthorizationBlocked`, or `Invalid`; only `AuthorizationBlocked` is an authorization stop.

Every governed object must have one unambiguous plan result. Evidence outside the current execution boundary must remain visible instead of disappearing from the report.

Conditional SharePoint capabilities are ingredients rather than hidden prerequisites. For example, a List content-type parent may require the Asset Library, Document Sets, or Video and Rich Media site feature. The plan records the feature ID, scope, dependency order, consuming content-type IDs, expected runtime content-type IDs, target probe, and activation action. A List collision does not make an independently activatable platform feature incompatible; the graph dependency keeps the consuming List gated while preserving the feature's own capability result.

## Target evidence model

The architecture distinguishes four target concepts:

| Concept | Meaning |
| --- | --- |
| Target mapping | The approved destination and transformation intent. |
| Target probe | What planning or fresh admission observed at that destination. |
| Target receipt | What execution actually created, reused, or recovered, including runtime IDs. |
| Verification result | Whether a fresh readback satisfies the approved expectation. |

A target URL in a plan is not proof that the Web exists. A target List ID in a receipt is not proof that its schema matches. The four concepts remain separate so that stale planning observations and incomplete mutations cannot masquerade as verified results.

Default topology mapping preserves the complete source path relative to its Site Collection root. The target Site Collection leaf may receive an isolation suffix, and a proven foreign collision may receive a stable suffix at the colliding node; Web, library, folder, and page tails otherwise remain byte-for-byte path-equivalent after URL decoding. Missing target nodes are materialization work, not mapping failures.

For Publishing Pages, target Content Type selection begins with the approved Page Layout association. Planning seals one exact Pages-library Content Type ID. Multiple descendants are an ambiguity blocker; Import and fresh readback require exact equality with the sealed ID rather than accepting any broad Enterprise Wiki descendant.

## Execution architecture

Before target mutation, import validates package structure, supported top-level schema versions, canonical digests, package state, the caller's approved plan digest, and fresh target admission. Nested contracts are structurally validated, but the current validators do not uniformly enforce every nested `schemaVersion`; this remains an explicit contract-hardening gap.

Malformed, unsupported, or digest-invalid packages currently throw before an operation ID or import receipt is created. Once contract validation succeeds, approval or fresh-target rejection returns a typed zero-mutation receipt.

Execution then follows dependency order. Each mutating category writes a `MigrationMutationIntent` before the SharePoint operation and a `MigrationMutationReceipt` after it completes or is proven already satisfied. Domain receipts expose runtime-generated target identities to later actions.

The execution journal is an audit and recovery primitive. It is not a rollback log. Safe retry is based on deterministic targets, ownership markers, semantic digests, and fresh target inspection.

See [Execution and verification](execution-and-verification.md) for detailed state semantics.

## Verification architecture

Verification is layered:

1. **Mutation outcome** records whether an attempted step was applied, already satisfied, or failed.
2. **Domain fresh readback** checks the supported closure for Topology, Lists, fields, Web Parts, content, content type, and lifecycle.
3. **Storage verification** aggregates required library-owned assertions.
4. **Runtime verification** is performed by an external browser-capable runner when required.
5. **Acceptance** combines the required storage and runtime outcomes.

The importer should fail storage verification when a required List or Topology assertion fails even if the page file itself was created successfully.

## Reporting architecture

The Markdown migration report is a human-review projection of the authoritative migration package. It can show complete bounded source evidence, every planned action, target probes, mappings, blockers, warnings, expected assertions, and both approval digests.

It cannot show actual post-import target identities or verification results because those do not exist until execution. A future execution report should be generated from the migration package plus import receipt rather than mutating the original plan report.

## Extension rule

A new object domain or newly supported behavior should be introduced as a complete vertical slice:

```text
source evidence
    + plan action/disposition
    + target probe
    + dependency ordering
    + ownership/retry rule
    + execution receipt
    + fresh verification
```

An implementation that adds only capture creates evidence without restoration. An implementation that adds only a writer creates unreviewed and unverifiable mutation. Both may be intentional intermediate steps, but their incomplete status must remain explicit.
