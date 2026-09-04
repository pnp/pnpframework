# Object lifecycle

> Status: Draft
> Implementation status: Current governed objects with explicit gaps
> Scope: Capture, planning, target analysis, execution, and verification

## Conceptual lifecycle

Every object governed by a migration workflow follows the same conceptual lifecycle even though domains use different CLR contracts and disposition enums:

```text
source identity and evidence
    -> action selection
    -> target mapping
    -> target probe and admission
    -> dependency-ordered mutation
    -> actual target identity receipt
    -> fresh-readback verification
```

This is not implemented as one universal `MigrationAction` base class. A Web, field value, content type, reference, and file have different identity and materialization rules. The commonality is a design contract, not an inheritance hierarchy.

## Governed object contract

A complete governed object should answer the following questions:

1. **Source identity:** Which source object is this, and at what ownership scope?
2. **Evidence:** What exact facts, values, schema, bytes, and diagnostics were captured?
3. **Action:** What will the current workflow do, and why?
4. **Target mapping:** Where or to which runtime object should it map?
5. **Target probe:** What target state justified the action?
6. **Dependencies:** Which target objects or identity maps must exist first?
7. **Ownership:** May the importer create, reuse, recover, or modify this target object?
8. **Mutation:** What exact supported state will be written?
9. **Receipt:** Which actual target IDs and outcomes were observed?
10. **Verification:** Which fresh-readback assertions prove the supported result?
11. **Unsupported evidence:** What remains preserved but intentionally not restored?

Not every captured evidence node becomes an independent mutation. For example, item field values are executed within an item operation, and a raw unsupported value remains evidence-only. However, every captured object governed by the current plan must be accounted for by an explicit result.

## Cross-domain matrix

| Object | Source evidence | Planned result | Target evidence | Receipt/runtime mapping | Fresh verification |
| --- | --- | --- | --- | --- | --- |
| Source ASPX artifact | File identity, exact bytes, Page directive, digest, availability | Preserve as evidence and transform into a target family shell, or block | Target path, runtime, layout, and Content Type plan | Target file identity | Source artifact digest remains valid; target page assertions pass |
| CLR runtime / profile / cohort | Page and layout declared types, Content Type/layout/field signals | Select one runtime adapter; retain multiple profile signals; include/exclude/unknown cohort result | Target runtime compatibility | No copied source runtime ID | Browser/runtime requirements and error-shell checks |
| Canonical ingredient | Node, ownership, evidence digest, required/optional edges | Preserve, transform, substitute, drop, delegate, or block | Per-action target identity and capability | Typed domain receipt supplies actual IDs | Per-action assertions plus dependency-closure evaluation |
| Site Collection / Web | IDs, parent, URL, title, template, configuration | Map existing host, create/reuse/recover child Web, or block | Existence, Site/Web/parent IDs, template, ownership properties | Source Web ID -> target Site/Web ID and actual disposition | Hierarchy, identity, URL, template/configuration, ownership digest |
| Platform feature | Feature ID, scope, dependency relationship inferred from captured runtime content-type parents | Reuse active feature, activate required feature, or block its own capability | Scope existence, active state, `ManageWeb`, and promised runtime content types | Mutation journal records activated/already-satisfied outcome | Feature remains active and every promised runtime content type is freshly available |
| Site field | GUID, schema, portable digest, role, taxonomy binding | Require target runtime, create/reuse owned, or block | Existing ID/type/schema and manageability | Target field at approved owner Web | Portable schema and required binding |
| Site content type | ID, parent, metadata, FieldLinks, required field closure | Create owned, reuse owned, or block | Parent availability, existing CT/metadata/links/fields | Target content type identity at owner Web | Parent, metadata, flags, required FieldLinks |
| List/library | Site/Web/List IDs, title/path, template, settings | Create owned, reuse owned, or allocate a stable suffix at the colliding List leaf | Web availability, path/title collision, template, ownership, permissions | Source List ID -> final target List ID/path and actual disposition | Identity, settings, ownership, final path, full supported closure |
| List field | ID, names, type, schema, lookup/taxonomy binding | Require runtime, create/reuse, map, evidence-only, or block | Runtime field compatibility and dependency availability | Target field and downstream value-writer behavior | Type/schema/binding and supported values |
| List content type | Source List CT ID, exact site parent, metadata, FieldLinks, order | Materialize from target site parent | Parent resolution and List CT membership | Source List CT ID -> generated target List CT ID | Parent, metadata, flags, links, explicit visible order |
| List item/folder/document | Source item ID, path, content type, typed/raw values, bytes | Materialize supported current state | Target schema and dependency identity maps | Source item ID -> target item ID; target file identity | Supported values, paths, content type, exact current bytes |
| Attachment | Parent item, file name, exact bytes | Materialize current attachment | Target parent item mapping | Target attachment under mapped item | Name, length, and SHA-256 |
| View | Source View ID, scope, URL, query, fields, paging, links/XML | Create/reuse public or page-bound View, skip personal, or block | Existing title/type/path and collision state | Source View ID -> target View ID | Supported query, fields, paging, and binding state |
| Page field value | Definition, typed value, raw evidence, capture status | Apply, handled/skip/evidence-only, mapping required, or block | Target field existence/type/writeability | Per-field import result | Persisted supported value |
| Page security | Inheritance state, role assignments, principals, and role-definition names | Preserve inherited security, delegate unique ACL replay, or block when inheritance is required | Target inheritance state | No principal mapping in the current implementation | Inherited target state is read back; unique source assignments remain snapshot evidence only |
| Page reference/resource | Source URL, scope, payload/artifact, availability | Preserve, rewrite, materialize, delegate, or block | Target path/existence and payload safety | Materialized dependency count and step receipt | Target resource/path and rewritten content assertion |
| Classic Web Part | Source ID, export XML, type, zone/order/hidden, List binding | Copy, rebind after dependencies, or block | Portability policy and target dependency mapping | Imported part/result using target runtime IDs | Export digest/definition, zone, order, hidden state |
| Page Layout | Identity, gallery metadata, exact ASPX, controls/zones/resources/schema | Reuse stock, create owned, reuse owned, or block | Existing bytes, association, resources, permissions, schema | Materialization steps and selected target layout | Exact owned bytes/resources/schema or approved stock reuse |
| Publishing Page | Site/Web/file identity, content, layout, lifecycle, fields, Web Parts | Create at the mapped path or allocate a stable suffix only at the colliding file leaf | Pages library, target content type/layout, preferred/final page path | Target file ID, List item ID, version, content type, ownership provenance | Content digest, ownership, fields, Web Parts, CT, lifecycle |

## Shared action semantics

Domain-specific enum names remain authoritative. For review, they fit into common semantic categories:

| Category | Meaning | Current examples |
| --- | --- | --- |
| Create | Materialize a new migration-owned object at an approved free target. | `CreateOwned`, `CreateOrReuseOwned`, `MaterializeAtTarget` |
| Reuse | Accept an existing approved host, runtime object, or exactly matching owned object. | `ReuseApprovedHost`, `ReuseOwned`, `RequireTargetRuntime`, `ReuseTargetStock` |
| Recover | Claim an interrupted creation only when deterministic description and ownership checks agree. | `RecoverInterruptedCreate` |
| Transform/map | Replace source runtime identity or authored location with reviewed target identity. | `MapLookup`, `MapTaxonomy`, `RewriteToTarget`, `RebindListAfterMaterialization` |
| Apply | Write a supported value or setting to an already selected target object. | `PageFieldDisposition.Apply`, lifecycle application |
| Preserve/skip | Retain evidence while intentionally making no current target mutation. | `EvidenceOnly`, `SkipPersonal`, `SkipReadOnly`, `SkipCalculated`, `PreserveExternal` |
| Delegate | Record that another capability owns the required work. | `PageReferenceDisposition.Delegate` |
| Defer/mitigate | Refuse the current transaction while retaining the object in the RCA, evidence, implementation, re-capture, and re-plan queue. | Domain-specific `Block` dispositions and plan blockers are projected to final ingredient `Defer` |
| Authorization stop | Stop only the affected branch when retained wire evidence proves literal HTTP 401/403. | Final ingredient `Block` and aggregate `AuthorizationBlocked` |

These categories are explanatory. They do not replace domain enums or authorize one domain to apply another domain's rules.

Every non-empty canonical ingredient also receives one semantic `IngredientDisposition`: `Preserve`, `Transform`, `Substitute`, `Drop`, `Delegate`, `Defer`, or `Block`. Domain planners may retain local `Block` enum values to mean that their current plan is unavailable, but final page orchestration projects those findings to nonterminal `Defer`. Only validated literal wire HTTP 401/403 evidence may produce final ingredient `Block`. Required graph edges are validated independently of the domain-specific execution order. A consumer transform may release a dependency only by naming it explicitly in `releasedDependencyIngredientIds`.

## Topology lifecycle

### Capture

`SourceSiteCollectionSnapshot` and `SourceWebSnapshot` preserve Site/Web IDs, parent Web identity, absolute and server-relative URLs, title, template, configuration, availability, and diagnostics.

### Planning

`TopologyPlan` maps the source hierarchy to deterministic target Site/Web locations. Its default mapping preserves the complete source Site-relative Web path; an isolation suffix is added only to the target Site Collection leaf, and a stable collision suffix belongs only on a proven colliding node. `TopologyTargetAnalysis` records one probe for every mapped Web and selects:

- `ReuseApprovedHost` for explicitly approved target hosts;
- `CreateOwned` for an available child-Web target;
- `ReuseOwned` for an exact migration-owned match;
- `RecoverInterruptedCreate` for a deterministic interrupted child-Web creation that can be safely claimed;
- `CreateMissing` at the run-orchestration layer for an absent planned Site/Web node; absence is not an unmapped URL;
- local `Block` for parent/template mismatch, insufficient evidence, or a post-approval target change; final page orchestration treats it as `Defer`, not an authorization stop. A foreign collision discovered during planning is resolved by suffixing only that Web leaf and retargeting only its source-graph descendants before the plan digest is sealed.

### Execution and verification

Execution creates or resolves parent-before-child. `TopologyMaterializationReceipt` records source-to-target Web identities and actual disposition. Fresh readback checks the complete mapped hierarchy and ownership/mapping digest.

Current explicit gap: `CreateTargetSite` is represented but tenant-scoped Site Collection creation is not executable.

## Site schema lifecycle

### Fields

Portable schema evidence removes target-specific storage slots and runtime identities before digest comparison. Field plans distinguish:

- `RequireTargetRuntime`: the target platform/template must provide a compatible field;
- `CreateOrReuseOwned`: materialize or accept an exact migration-owned field with the source GUID;
- local `Block`: required schema cannot currently be satisfied; final orchestration emits `Defer` and keeps it in the mitigation queue.

Taxonomy field schema requires reviewed target Term Store/Term Set mapping. Source WssId is never a schema identity.

### Content types

Capture recursively records the minimal required custom ancestor and field closure. Runtime classification uses exact known content type IDs instead of treating every descendant of Item or Document as target runtime.

Plans select `CreateOwned`, `ReuseOwned`, or local `Block`. Target probes record parent availability, existing identity and metadata, FieldLinks, fields, collisions, permissions, and diagnostics. Local `Block` becomes final ingredient `Defer`. Execution orders ancestors before descendants and verifies parent, metadata flags, and required FieldLinks after materialization.

## List lifecycle

### Capture

`ListDependencySnapshot` preserves identity, owner Web, root folder, title, template, settings, all returned field evidence, custom site CT closure, List-local content types and explicit order, Views, current items/folders/files/attachments, and diagnostics. Lookup edges are captured separately for deterministic ordering and cycle detection.

For each returned List-item `FieldValues` entry, capture creates a `ListItemValueSnapshot`. Known runtime types receive a typed representation. Unknown types retain runtime type, best-effort invariant string, and best-effort raw JSON; if a lossless typed representation is unavailable, availability is `Partial` and diagnostics explain the limitation. The current planner writes only reviewed supported field kinds, while the raw snapshot remains available for a future materializer.

Binary capture distinguishes an ordinary file payload from an Information Rights Management envelope. In both cases the artifact SHA-256 seals the exact bytes SharePoint returned. For an IRM envelope, capture additionally retains the SharePoint `cTag` and `QuickXorHash` logical-content identity when present in `MetaInfo`. The envelope can be larger than the logical file and can receive a different byte SHA on every CSOM or REST read even when source identity, version, modified time, length, `cTag`, and `QuickXorHash` are unchanged. That difference is recorded as `RightsManagedEnvelopeLengthMismatch`, not as incomplete byte capture.

An immutable export written before this classifier existed deserializes as `Unclassified`; the canonical serializer omits that default value so the historical snapshot digest remains valid. The exact historical bytes stay sealed, but the loader does not invent ordinary-file semantics. Its binary ingredient is `Defer` with `ListBinaryRepresentationUnclassified` until a fresh source capture records an explicit classification.

### Planning

A `ListMaterializationPlan` selects `CreateOwned`, `ReuseOwned`, or local `Block`, plus nested field, View, and site-content-type plans. A local `Block` is projected to final ingredient `Defer`. Before sealing, a foreign path or title collision is resolved by allocating a stable suffix only at the List/library leaf; the preferred and final path/title plus the reason remain in `ListTargetProbe`. Apply performs a strict fresh probe and rejects a newly occupied sealed path rather than choosing another path after approval.

List field dispositions distinguish platform/runtime requirements, schema-only creation, value-copy creation, calculated dependency ordering, lookup mapping, taxonomy mapping, evidence-only retention, and blockers.

View dispositions distinguish public Views, deterministic page-bound Views, personal evidence-only Views, and blockers.

An IRM-envelope document currently receives ingredient `Defer` with mitigation code `ListRightsManagedBinaryReplayUnverified`. The exact response artifact remains in the snapshot, but it is not sent through the ordinary exact-byte document materializer until cross-site usability and a semantic target verifier are proven. This is nonterminal mitigation work and is not an authorization block.

### Execution and verification

Lists execute in lookup topological order. Before site content-type membership or List creation, the importer activates required conditional platform features in dependency order and verifies their promised runtime content types. It then materializes site CT closure, List identity/settings, fields, List CTs and order, folders/items/files/attachments, and Views. It records target List, item, View, and List-local CT IDs for downstream lookup values and Web Part rebinding.

Fresh verification checks the supported List closure and exact current file/attachment bytes. Rights-managed documents require a future semantic verifier over logical content identity because a fresh download envelope is not byte-stable. A List mismatch makes `listsMatched=false` even when the page was created.

## Page field lifecycle

Capture retains every field returned for the source Pages-library item, including raw fallback evidence. Planning creates exactly one `PageFieldAction` per captured field.

Current dispositions are:

- `Apply`;
- `AlreadyHandled`;
- `SkipEmpty`;
- `SkipReadOnly`;
- `SkipCalculated`;
- `TargetFieldMissing`;
- `TargetTypeMismatch`;
- `RequiresMapping`;
- `EvidenceOnly`;
- `CaptureUnavailable`;
- `Block`.

Only `Apply` authorizes a field write. Other results remain visible in the package/report. A domain `Block` contributes final ingredient `Defer` unless validated literal HTTP 401/403 evidence exists for that ingredient.

## Reference and Web Part lifecycle

References may be preserved externally, rewritten, materialized from captured payload, delegated, or locally blocked pending mitigation. Replacements are explicit plan content and participate in `planDigest`; local blockers project to final ingredient `Defer`.

Classic Web Parts are captured as shared evidence. `ClassicWebPartReplayCapabilityPolicy` and `ClassicWebPartActionPlanner` choose `CopyCaptured`, `RebindListAfterMaterialization`, or local `Block` for the current Publishing Page importer. Local `Block` becomes final ingredient `Defer`. A List-bound Web Part executes only after target Web/List/View identity maps exist. Verification reads back supported definition and placement properties.

## Page and lifecycle result

The current executable page operation is `CreatePage`; `createOnly` prevents overwriting an existing target page. Target lifecycle is derived rather than supplied by a top-level publish Boolean:

- source evidence that is consistently published may produce `Published`;
- all other supported source lifecycle states produce `Draft`.

The receipt records actual file level, checkout type, moderation status, content type, version, publishing-content digest, field results, and Web Part results. Storage verification passes only when all required page and dependency assertions agree.

## Completeness rule

A future object is not fully supported merely because it can be captured or written. It becomes a governed executable object only when all lifecycle stages are defined and the package/report make the result reviewable.
