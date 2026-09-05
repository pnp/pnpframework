# Page classification and ingredient policy

> Status: Draft
> Implementation status: Implemented for the Publishing Page / Enterprise Wiki v1 workflow
> Scope: `PnP.Framework.Migration.Pages`

## Why page type is layered

A SharePoint page cannot be classified safely by one label. The migration model uses four separate layers:

1. **CLR runtime adapter** answers which executable page runtime renders the ASPX file.
2. **Profile signals** describe non-exclusive product ancestry and traits such as Content Type lineage, Page Layout, and known fields.
3. **Validation cohort** answers whether the page belongs to the population reviewed by one workflow version.
4. **Ingredient graph** describes the actual page plus external content that must be assessed independently.

These layers must not be collapsed. A Project Page can descend from Enterprise Wiki Page and therefore emit both Enterprise Wiki and Project Page profile signals, while the EW-v1 validation cohort deliberately excludes it. Likewise, an ASPX file whose `Page` directive declares `WikiEditPage` selects the Wiki runtime even if its List item has an Enterprise Wiki Content Type.

## Runtime selection

`PageRuntimeResolver` uses this priority:

1. source ASPX `Page` directive CLR type;
2. Page Layout `Page` directive CLR type;
3. known Content Type lineage only as an explicit fallback;
4. `runtime.unknown` when no executable adapter can be established.

Current adapter IDs are:

| Adapter | Meaning |
| --- | --- |
| `runtime.publishing` | Classic Publishing Page runtime. |
| `runtime.wiki` | Wiki edit-page runtime. |
| `runtime.webpart-page` | Classic Web Part Page runtime. |
| `runtime.unknown` | No safe executable adapter was resolved. |

The source ASPX bytes and parsed directive are retained in `snapshot.pageArtifact`; runtime selection is therefore reviewable and can be recomputed from evidence.

## Profiles and cohorts

`snapshot.profileSignals` is a list rather than a single discriminator. Signals currently include:

- Content Type ancestry;
- selected Page Layout filename;
- presence of known page fields.

Signals are facts, not authorization and not action selection. `selection.validationCohort` is a workflow-owned assessment with its own ID, policy version, disposition, and reasons. `selectionDigest` seals that choice, and the selected planner/importer recomputes it from source evidence before use. The first workflow is `enterprise-wiki-v1`.

Cohort membership and migration capability are intentionally different questions. A page may be technically reproducible but outside a validation cohort, or inside the cohort but blocked by one unsupported ingredient.

## Canonical ingredient graph

`snapshot.ingredientGraph` projects the typed snapshots into `CanonicalPageIngredientGraph`. Each node records:

- stable ingredient ID and kind;
- whether meaningful content is present;
- ownership boundary;
- authoritative source evidence and optional digest;
- required runtime;
- references back to detailed typed evidence.

Edges record consumer-to-dependency relationships, their semantic kind, and whether the dependency is required, conditional, or optional. The graph does not replace typed snapshots. It is an index and dependency model over them.

Current Publishing Page ingredients are projected at the smallest independently planned object boundary:

| Domain | Ingredient nodes |
| --- | --- |
| Page core | CLR runtime, source ASPX artifact, `PublishingPageContent`, page-item fields, security, and lifecycle. |
| Publishing layout | Page Layout, associated Content Type, each required associated site-field schema, and each authored layout resource. |
| Topology | Every captured owner/ancestor Web, with child-to-parent and object-to-owner edges. |
| List schema | Each List/library, site Content Type in its required parent closure, required site field, List-local Content Type, and List field. |
| List content | Each current List item, document/folder, attachment, and View. Item-to-field and item-to-Content-Type edges expose the schema needed by actual values. |
| Page composition | Each classic Web Part and authored reference, including required Web/List/View bindings. |

The top-level graph and action projectors are only coordinators. Core, layout, topology, List schema, List content, Web Part, and reference projectors own their corresponding projection rules. This keeps Enterprise Wiki as a workflow facade and prevents a new all-purpose page class from absorbing unrelated object domains.

## Ingredient action contract

Planning creates exactly one `PageIngredientAction` for every non-empty ingredient. An action records:

| Field | Meaning |
| --- | --- |
| `actionId` | Stable action identity inside the plan. |
| `ingredientId` | Captured graph node governed by the action. |
| `capability` | Whether the selected target can represent the ingredient. |
| `disposition` | Semantic decision: preserve, transform, substitute, drop, delegate, defer, or authorization-block. |
| `realization` | Concrete implementation mechanism. |
| `targetIdentity` | Planned target path, field, Content Type, runtime, or other locator. |
| `policyId` / `policyVersion` | Rule that made the decision. |
| `reason` | Human-readable rationale. |
| `releasedDependencyIngredientIds` | Required dependencies deliberately eliminated by this ingredient's transform. |
| `verificationAssertions` | Ingredient-specific expectations that must be checked by storage or runtime verification. |

Disposition semantics are:

| Disposition | Interpretation |
| --- | --- |
| `Preserve` | Retain the ingredient's relevant semantics at the target. |
| `Transform` | Deliberately change representation while retaining the reviewed intent. |
| `Substitute` | Use a target-runtime supplied equivalent. |
| `Drop` | Deliberately omit the ingredient and record reviewed loss. |
| `Delegate` | Keep source evidence but assign restoration to another workflow. |
| `Defer` | Keep a known evidence, mapping, or capability gap in the nonterminal mitigation and re-planning queue. |
| `Block` | Stop only this ingredient branch because the package retains digest-valid literal wire HTTP 401/403 evidence for it. |

The source ASPX artifact is a useful example: the Publishing workflow does not deploy the source file as executable code. It creates a target Publishing Page shell and retains the exact source bytes as evidence, so this action is `Transform`, not `Preserve`.

List item capture and List field policy are intentionally asymmetric. Capture keeps every returned item field value, including typed values and best-effort raw type/text/JSON for unknown CLR values. Planning writes only fields with an understood materializer:

| List field plan | Ingredient result |
| --- | --- |
| target-runtime field | `Substitute`; SharePoint owns the target schema/value where applicable. |
| source-owned supported field | `Preserve`; create or exactly reuse the schema and copy recognized values. |
| lookup or taxonomy field | `Transform`; rewrite site-local identities through reviewed target mappings. |
| unsupported field with no nonempty item, View, or Content Type consumer | `Drop`; retain its complete snapshot for later recovery. |
| unsupported field used by captured content or schema | `Defer`; omission would break a retained consumer, so continue RCA/materializer work and re-plan. |

## Dependency closure and discard rules

A retained consumer cannot silently lose a required dependency. `PageIngredientPlanEvaluator` marks the action graph invalid when:

- a non-empty ingredient has no action;
- an action has an undefined disposition, or uses `Block` without matching literal HTTP 401/403 evidence;
- a retained ingredient has unknown capability;
- a required dependency is dropped, delegated, or missing.

A required dependency in `Defer` makes the aggregate `MitigationPending`; a required dependency in evidence-backed `Block` makes the affected branch `AuthorizationBlocked`.

A dependency may be discarded only when one of these conditions holds:

1. no retained consumer requires it;
2. the consumer itself is dropped or delegated;
3. a consumer `Transform` explicitly lists the dependency in `releasedDependencyIngredientIds` and the transform's rationale and verification make the replacement behavior reviewable.

The evaluator rejects a release list when the action is not `Transform`, the released ID is missing or duplicated, or the named node is not an actual required outgoing dependency of that consumer. Merely adding an ID to the list cannot bypass dependency closure.

Releasing a dependency does not make the result exact. A `Drop` or `Delegate` produces `ExecutableWithLoss`; a `Transform` or `Substitute` produces at least `ExecutableWithTransform`.

For an admitted reviewed stock Page Layout, embedded stock CSS, image, and target-runtime references are `Substitute` ingredients rather than copied assets. The exact target stock layout readback proves which references are owned by that target layout; runtime verification must still prove that they resolve. Page-item values remain separate decisions: a SharePoint-owned source field becomes `TargetRuntime` only when the target exposes an equivalent same-name, same-type field, and then projects as a target-runtime substitution. `FileLeafRef` is handled by target page-shell creation. Unknown or unmatched business-field values remain delegated evidence and are not upgraded by this rule.

## Aggregate outcome

The evaluator derives one `plan.migrationOutcome`:

| Outcome | Meaning |
| --- | --- |
| `Exact` | All non-empty ingredients are preserved without transform, substitution, delegation, or loss. |
| `ExecutableWithTransform` | The plan is executable but at least one ingredient changes representation. |
| `ExecutableWithLoss` | The plan is executable but at least one ingredient is deliberately dropped or delegated. |
| `MitigationPending` | At least one ingredient is `Defer`; retain it in the nonterminal mitigation queue. |
| `AuthorizationBlocked` | At least one branch is `Block`, backed by retained literal wire HTTP 401/403 evidence. |
| `Invalid` | The proposed action graph is inconsistent and returns to RCA; it is not an authorization stop. |
| `Unknown` | No trustworthy evaluation exists; this is never importable. |

`plan.isExecutable` additionally requires the workflow-wide blocker list to be empty. The ingredient outcome does not hide target topology, layout admission, or policy blockers that live in typed domain plans. A foreign Web/List/Page collision observed before sealing is a deterministic retargeting decision at that node, not a terminal workflow blocker; a collision observed after approval is a stale-plan precondition failure that requires replanning.

## Realistic Enterprise Wiki example

Consider an Enterprise Wiki page with a stock layout, one List View Web Part, an authored same-site URL, an unknown custom field, inherited permissions, and a draft source version:

| Ingredient | Capability / disposition | Target and verification |
| --- | --- | --- |
| CLR runtime | available / preserve | `runtime.publishing`; page must load without an error shell. |
| Source ASPX artifact | available / transform | Create `/sites/target/Pages/source.aspx`; retain and digest-check source bytes. |
| Page Layout | available / preserve | Reuse approved `EnterpriseWiki.aspx`; verify layout association. |
| Page Content Type | available / preserve | Seal one exact Pages-library Content Type ID and require exact post-create equality. |
| Publishing content | available / transform | Rewrite only approved URLs and verify the expected content digest. |
| Dependency List and View | available / preserve | Create or reuse owned objects and verify schema plus generated runtime IDs. |
| List View Web Part | available / preserve | Rewrite Web/List/View IDs from receipts and verify placement/binding. |
| Unknown nonempty page field | unknown / delegate | Keep schema and raw value in the snapshot; do not guess a write. |
| Unused unsupported List field | unknown / drop | Omit it only because no retained item, View, or Content Type requires it; retain the full field/value snapshot. |
| Security | available / preserve | Reuse target inheritance and verify that the page inherits permissions. |
| Lifecycle | available / preserve | Keep the target as Draft and verify the final file state. |

The plan is executable, but its aggregate outcome is `ExecutableWithLoss` because the unknown field is delegated. A later mapper can recover that field from the original snapshot without recapturing the source.

## Transaction analogy

Each typed domain action is transaction-like: it has sealed intent, target preconditions, a mutation intent/receipt, and fresh verification. The whole migration is not an atomic database transaction. SharePoint can allocate runtime IDs and partial mutations can occur, so recovery relies on deterministic ownership, content digests, fresh probes, and journals rather than global rollback.
