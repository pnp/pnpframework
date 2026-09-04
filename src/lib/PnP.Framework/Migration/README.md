# Migration

The `PnP.Framework.Migration` namespace contains staged, evidence-preserving migration workflows for SharePoint artifacts.

Migration is intentionally separate from the existing provisioning and modernization engines:

| Area | Primary purpose |
| --- | --- |
| Provisioning | Describe or apply a desired SharePoint configuration. |
| Modernization | Transform supported source experiences into modern SharePoint experiences. |
| Migration | Capture source evidence, make target-specific decisions, apply only an approved plan, and verify the result from a fresh readback. |

The migration APIs are being introduced incrementally. A namespace or contract is not considered stable merely because it is public; schema versions and release notes define compatibility once a feature is released.

## Design documentation

Contributor-facing subsystem design documents live in [`docs`](docs/README.md). They describe the purpose and scope, architecture, package model, per-object lifecycle, and execution/verification semantics. These code-adjacent drafts are intentionally separate from the repository-root DocFX customer documentation.

Start with:

- [Purpose and scope](docs/purpose-and-scope.md)
- [Architecture](docs/architecture.md)
- [Package model](docs/package-model.md)
- [Page classification and ingredient policy](docs/page-classification-and-ingredient-policy.md)
- [Object lifecycle](docs/object-lifecycle.md)
- [Execution and verification](docs/execution-and-verification.md)
- [Performance and concurrency](docs/performance-and-concurrency.md)

## Design rules

Every migration area should preserve the following boundaries:

1. **Capture is source-only.** Export must not require a target connection and must not make target decisions.
2. **Planning is target-specific and read-only.** A planner inspects the target, records every decision, and produces blockers and warnings without changing the target.
3. **Import executes a sealed plan.** The importer validates both the source snapshot digest and the approved plan digest before writing.
4. **Verification uses a fresh readback.** Success is based on persisted target state, not only on successful CSOM requests.
5. **Unknown evidence is retained.** Capture should preserve information even when the current importer cannot restore it. Planning decides what is understood and safe to apply.
6. **Unsafe ambiguity becomes a blocker or a conservative result.** A migration workflow must not silently guess cross-site identities, target bindings, runtime adapter, required dependencies, or lifecycle state.
7. **Runtime, profiles, cohorts, and ingredients are separate.** CLR evidence selects an executable adapter; profile signals are non-exclusive; cohort membership is workflow-specific; every non-empty ingredient receives an explicit action.
8. **Profiles compose reusable page capabilities.** Cross-page evidence and mechanics belong in `Pages`; publishing-page contracts and lifecycle behavior belong in `Pages.Publishing`; the Enterprise Wiki namespace remains a thin workflow facade.
9. **Use existing PnP primitives.** Migration code should compose established PnP Framework operations instead of duplicating CSOM retry, file, folder, page, URL, or Web Part plumbing.
10. **Optimize round trips without weakening evidence.** Batch compatible target reads, reuse already-loaded properties inside one inspection boundary, and prove that optimization leaves canonical plan digests and typed decisions unchanged.

## Current areas

| Namespace | Scope |
| --- | --- |
| `PnP.Framework.Migration.Diagnostics` | Typed, stable migration issues that can be reported without parsing exception or blocker text. |
| `PnP.Framework.Migration.Evidence` | Evidence availability, source lineage, and derived-artifact provenance shared by migration domains. |
| `PnP.Framework.Migration.Execution` | Operation state, write-ahead mutation intents, step receipts, and pluggable execution journals. |
| `PnP.Framework.Migration.Packaging` | Content-addressed artifact references, artifact-store contracts, digest helpers, and a local directory-backed content-addressed store for larger or binary evidence. |
| `PnP.Framework.Migration.Topology` | Source SPSite/SPWeb evidence, complete Site-relative target maps, collision probes, migration-owned provenance, child-Web materialization, and source-to-target runtime identity receipts. |
| `PnP.Framework.Migration.Features` | Explicit conditional SharePoint platform-feature plans, target probes, dependency-ordered activation, and promised runtime-contract verification. |
| `PnP.Framework.Migration.Lists` | Page-required List/library dependency closure, lookup ordering, target planning, create-or-owned-reuse execution, and final fresh-readback results. |
| `PnP.Framework.Migration.Lists.Fields` | Complete List field-schema evidence and List-specific schema/value planning. |
| `PnP.Framework.Migration.Lists.ContentTypes` | List-local content types and exact FieldLink evidence. |
| `PnP.Framework.Migration.Lists.Items` | Complete current item value evidence plus folders, current file bytes, and attachments. |
| `PnP.Framework.Migration.Lists.Views` | Public, embedded/page-bound, and personal View evidence; personal Views remain evidence-only. |
| `PnP.Framework.Migration.Schema.Fields` | Portable field-schema evidence, ownership classification, canonicalization, exact-ID materialization plans, and target probes. |
| `PnP.Framework.Migration.Schema.ContentTypes` | Minimal required-field content-type closure capture, planning, target admission, exact-ID materialization, and fresh verification. |
| `PnP.Framework.Migration.Taxonomy` | Exact taxonomy relationship evidence, reviewed mappings, source TermSet/Term asset closure, migration provenance, and conservative target classification. |
| `PnP.Framework.Migration.Verification` | Storage/runtime verification states and typed external runtime-verification manifests and receipts. |
| [`PnP.Framework.Migration.Pages`](Pages/README.md) | Shared page identity, exact ASPX evidence, CLR runtime, non-exclusive profiles, cohorts, canonical ingredients, fields, references, security, classic Web Parts, content, capture, and planning capabilities. |
| [`PnP.Framework.Migration.Pages.Publishing`](Pages/Publishing/README.md) | Classic publishing-page aggregate contracts, workflow policy, lifecycle, planning, packages, reports, execution, and verification. |
| `PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki` | Thin Enterprise Wiki v1 discovery, API, and EW-named file-storage facade. |

Future Wiki Page and Web Part Page implementations should be sibling page families under `Pages`, reusing the shared page capabilities instead of copying publishing-page types.

## Cross-site object convergence

The first implementation was informed by the parallel Repro4 proof of concept, but PnP Framework does not reference, execute, or mutate Repro4. Repro4 remains a moving validation workspace. A behavior is absorbed here only after it can be expressed as a reusable object-domain contract, sealed plan decision, conservative target admission rule, and fresh-readback assertion.

The dependency direction is:

```text
source evidence
    -> CLR runtime + non-exclusive profile signals + canonical ingredient graph
    -> Topology plan (SPSite/SPWeb ownership)
    -> conditional platform-feature requirements
    -> shared Schema closure (site fields/content types)
    -> List closure (List, list CTs, fields, items/files, views)
    -> Classic Web Part runtime rebinding
    -> page-family write and lifecycle
    -> storage receipt and optional external runtime acceptance
```

The object model deliberately does not place these capabilities under an `EnterpriseWiki` namespace. Enterprise Wiki is a page-profile entry point. `Topology`, `Features`, `Schema`, `Lists`, `Taxonomy`, and shared `Pages.ClassicWebParts` own their respective evidence and mechanics so future Wiki Page and Web Part Page profiles can compose the same implementation.

### Absorbed proof-of-concept assets

| Proven behavior | PnP Framework expression |
| --- | --- |
| Preserve SPSite versus SPWeb identity level and child-Web ancestry. | `SourceSiteCollectionSnapshot`, `TopologyPlan`, `TopologyTargetAnalysis`, and `TopologyMaterializationReceipt`. |
| Preserve the complete Site/Web/Library/Folder/Page relative path and never overwrite an unowned collision. | Planning changes no relative segment. A stable suffix is added only at the topology/object node used for run isolation or where an observed ownership collision requires it, then per-object original-identifier plus semantic-digest provenance is sealed; Apply rejects any post-approval path change and requires replanning. |
| Capture the complete List/lookup closure required by a page. | `ListDependencySnapshot`, `ListLookupDependency`, and deterministic DAG ordering with cycle blocking. |
| Preserve unknown item values for future recovery while writing only understood fields. | Every returned `ListItemValueSnapshot` keeps typed and raw evidence; `ListFieldMaterializationDisposition` controls replay. |
| Recreate custom site-content-type ancestry without treating every child of Document as runtime. | Exact runtime content-type catalog plus `ContentTypeClosureSnapshotReader`, planner, and materializer. |
| Do not pollute business content types with helper fields. | Migration-owned List fields use `AddToNoContentType`; FieldLinks are applied explicitly. |
| Preserve List-local content-type shape and order. | Parent-based source-to-target List CT ID mapping, metadata and FieldLink replay, null-versus-explicit order evidence, filtering of disallowed Folder/UntypedDocument children, and exact readback. |
| Create calculated fields after their calculated dependencies. | Formula-reference dependency ordering, including display-name references, with cycle blocking. |
| Accept runtime field evolution only inside a known serialized-value family. | Shared compatibility rules allow equivalent scalar representations such as Text/Note/Choice and numeric types while keeping every single-value versus multi-value shape distinct. |
| Never force source List/View/item/WssId values into the target. | Runtime-generated IDs are recorded in List receipts; lookup and classic Web Part consumers are rewritten through those maps; taxonomy writes ignore source WssId. |
| Preserve taxonomy identity without healing invalid relationships. | `Taxonomy.Assets` captures only the required TermSet/Term/ancestor closure, derives an explicit deterministic target TermGroup ingredient, preserves exact GUIDs and Repro4-compatible original-identifier provenance, classifies owned/external/missing/colliding target assets, and keeps page relationship replay independent from asset preparation. |
| Prove more than successful mutation calls. | Final topology and per-List fresh readback covers identity, settings, provenance, schema, CT metadata/order/FieldLinks, supported Views, current item values, files, and attachments. |

### Deliberately open boundaries

The following proof-of-concept behaviors are not silently approximated:

- Site-collection creation needs a tenant-scoped executor; the current importer accepts an existing target site collection and can create/recover mapped child Webs only.
- Child-Web feature activation is not inferred from the source template. The generic materializer uses the sealed target Web template; Publishing/Enterprise Wiki, Document ID, Document Set, asset-library, and other Feature prerequisites still need explicit capability plans.
- A same-title template-created List is a blocker. `ListTargetOverride.TargetTitle` supports a reviewed alternate target title, but PnP does not yet rename a selected template List through the proof-of-concept's resumable claim protocol.
- List View/Web Part `JSLink` and `XslLink` strings are captured, but custom referenced bytes do not yet have a List-rendering-resource artifact/materialization contract. Custom paths block planning.
- Taxonomy asset capture, deterministic planning, read-only target inspection, digest-sealed ingredient-level approval, journaled materialization, and aggregate fresh verification are implemented under `Migration.Taxonomy.Assets`. TermGroup, TermSet, and Term actions are explicit and dependency-checked; no child approval authorizes hidden parent creation. The 10% plan has not been approved or applied to the live target. Mutating an external TermSet requires a separate per-Term authorization, and retaining many-to-one source aliases remains follow-up work.
- Full `ListViewXml`, View hidden/default repair in every collision shape, exact removal of extra List CT FieldLinks, and template/Feature-specific List creation remain narrower follow-up work.
- Version history, audit identity/timestamps, unique ACLs, workflows, subscriptions, event receivers, personal Views, and browser DOM/visual acceptance remain outside the storage importer.

These are named admission gaps. A future change should add evidence, a disposition, target probing, execution ownership, and verification together; it should not remove a blocker in isolation.

## Namespace and folder ownership

Folders represent an object domain or a public workflow boundary, not a generic implementation role. Avoid broad folders such as `Helpers`, `Managers`, `Readers`, or `Writers` when the type naturally belongs to a domain such as `Fields`, `WebParts`, or `Lifecycle`.

A profile may coordinate several domains, but it should not absorb shared evidence models or mechanics. Conversely, shared page code must not contain publishing layout assumptions, Enterprise Wiki content type IDs, portability decisions, or target-template policy. Dependencies point inward: profile -> page family -> shared page capabilities.

## Adding a migration area

Before adding another migration area or profile:

- define the source identity and stability fence;
- define a versioned export contract that retains unsupported evidence;
- define target probes and explicit per-object actions;
- define the plan digest that represents the review boundary;
- document blockers, warnings, and conservative defaults;
- define fresh-readback evidence and an import receipt;
- add mutation tests proving that snapshot and plan changes invalidate their digests;
- add contract round-trip tests and focused policy tests;
- document which existing PnP Framework primitives are reused.

Do not add a direct source-to-target copy path that bypasses export, planning, approval, or verification.

## Execution and runtime acceptance

An importer records `NotStarted`, `Running`, `Succeeded`, or `FailedUnexpectedly` independently from source eligibility and plan approval. Expected admission failures return a zero-mutation receipt. Once execution starts, each mutating step writes an intent before the SharePoint operation and a receipt after the operation returns.

Fresh storage verification and browser/runtime acceptance are also separate. PnP Framework owns storage readback. A browser-capable external runner consumes a typed `RuntimeVerificationManifest` and returns a digest-bound `RuntimeVerificationReceipt`; recording a requirement does not imply that it ran.
