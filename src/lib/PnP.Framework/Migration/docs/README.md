# Migration design documentation

> Status: Draft
> Scope: `PnP.Framework.Migration`
> Audience: Contributors and reviewers
> Publication: Code-adjacent documentation; not included in the PnP Framework DocFX site

This directory is the design record for the staged migration subsystem. It explains why the subsystem exists, how its artifacts relate, how a captured object becomes a target-specific action, and what execution and verification mean.

The documents describe the current implementation and its intended invariants. They are not a customer-facing usage guide and do not make compatibility promises beyond the versioned contracts implemented in code.

## Reading order

1. [Purpose and scope](purpose-and-scope.md) explains the problem, goals, and non-goals.
2. [Architecture](architecture.md) describes the stages, dependency direction, and domain boundaries.
3. [Package model](package-model.md) describes the export package, migration package, plan, import receipt, and digest relationships.
4. [Object lifecycle](object-lifecycle.md) follows each governed object from captured evidence through action selection, target resolution, execution, and verification.
5. [Execution and verification](execution-and-verification.md) defines admission, journaling, retry, receipt, fresh-readback, and acceptance semantics.
6. [Page classification and ingredient policy](page-classification-and-ingredient-policy.md) defines CLR runtime selection, non-exclusive profiles, validation cohorts, ingredient actions, dependency release, and aggregate outcomes.
7. [Taxonomy relationship fidelity](taxonomy-relationship-fidelity.md) defines exact capture and reproduction of valid, outside-bound, and dangling taxonomy relationships without Term repair.
8. [Performance and concurrency](performance-and-concurrency.md) defines measurement stages, safe read parallelism, write serialization, caching boundaries, and semantic regression requirements.

## Document roles

These design documents and the code-adjacent README files have different roles:

| Location | Role |
| --- | --- |
| `Migration/README.md` | Short subsystem entry point, design rules, and namespace index. |
| `Migration/Pages/README.md` | Shared page-kernel ownership and dependency rules. |
| `Migration/Pages/Publishing/README.md` | Publishing Page family contract and current implementation details. |
| `Migration/docs/` | Cross-domain design, artifact semantics, rationale, and contributor guidance. |
| Repository-root `docs/` | Future customer-facing DocFX documentation. It is intentionally out of scope for this draft. |

Code is authoritative when a document and implementation disagree. A contract change should update the relevant design document in the same pull request so that disagreement is temporary and visible.

## Normative language

The words **must**, **must not**, **should**, and **may** describe design constraints:

- **must** and **must not** identify invariants required for safety or contract integrity;
- **should** identifies the expected design unless a narrower domain documents a reason to differ;
- **may** identifies an optional behavior.

An item labelled **Implemented** exists in the current code. **Planned** identifies an intended extension that is not an import capability yet. **Explicit gap** identifies behavior that must remain blocked, evidence-only, or externally verified until its complete lifecycle is implemented.

## Current versioned artifacts

| Artifact | Schema |
| --- | --- |
| Publishing Page source export | `pnp-publishing-page-export/v2` |
| Publishing Page migration package | `pnp-publishing-page-migration-package/v2` |
| Publishing Page import receipt | `pnp-publishing-page-import-receipt/v2` |
| Source ASPX artifact | `pnp-page-artifact/v1` |
| Page runtime resolution | `pnp-page-runtime/v1` |
| Canonical page ingredient graph | `pnp-page-ingredient-graph/v1` |
| Source topology | `pnp-source-topology/v1` |
| Topology plan | `pnp-topology-plan/v1` |
| Topology target analysis | `pnp-topology-target-analysis/v1` |
| List dependency snapshot | `pnp-list-dependency/v1` |
| List migration plan | `pnp-list-migration-plan/v1` |
| Content type schema snapshot | `pnp-content-type-schema/v1` |
| Publishing Page Layout snapshot | `pnp-publishing-page-layout/v1` |
| External artifact manifest | `pnp-migration-artifacts/v1` |
| Runtime verification manifest | `pnp-migration-runtime-verification/v1` |
| Runtime verification receipt | `pnp-migration-runtime-verification-receipt/v1` |
| Taxonomy value relationship | `pnp-taxonomy-value-relationship/v1` |

These contracts are under development. A CLR type being public does not by itself declare the contract stable or released.

## Updating these documents

A change that adds a governed object or removes a blocker should describe all of the following together:

1. source identity and captured evidence;
2. action or disposition choices;
3. target mapping and target probe;
4. dependency ordering;
5. ownership and retry semantics;
6. execution receipt and runtime identity mapping;
7. fresh-readback assertions;
8. evidence that remains unsupported.

Adding only a writer is insufficient. A migration capability is complete only when its capture, planning, execution, and verification semantics agree.
