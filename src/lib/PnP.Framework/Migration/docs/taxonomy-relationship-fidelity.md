# Taxonomy relationship fidelity

> Status: Draft
> Implementation status: Implemented for Publishing Page fields
> Applies to: page-field taxonomy capture, planning, import, and fresh verification

## Decision

Migration preserves the relationship observed at the source. It does not repair taxonomy.

An invalid source relationship is therefore not treated as a request to create a missing Term, move a Term into the field's bound TermSet, substitute a same-label Term, or silently drop the value. If the owning field is selected for replay and the exact relationship can be proven and reproduced, the plan contains an explicit relationship action. If a selected relationship cannot be reproduced, that ingredient blocks.

This rule applies independently to every taxonomy value. The page field is the consumer; each value relationship is a required taxonomy ingredient in the canonical dependency graph.

Capture scope remains wider than restore scope. When the current workflow does not select a taxonomy field for replay, every captured value receives `RetainEvidenceOnly`. That action preserves the sealed relationship proof for a later workflow, makes no target capability claim, performs no target taxonomy probe, and projects to the ingredient-level `Delegate` disposition. Source conflict evidence in an unselected field therefore remains recoverable without blocking unrelated page migration.

## Source relationship states

| State | Source fact | Selected-field target interpretation |
| --- | --- | --- |
| `LiveInBoundTermSet` | The Term GUID resolves live inside the field's bound TermSet. | Reuse the same GUID in the explicitly mapped target bound TermSet. The target Term must already exist with the exact captured label and path. |
| `LiveOutsideBoundTermSet` | The GUID is absent from the bound TermSet but resolves as a live global Term in another TermSet. The page's `TaxCatchAll` retains an `UNVALIDATED` row for the bound TermSet. | Reproduce the same outside-bound relationship. Both source TermSets require reviewed mappings, the exact live Term must already exist outside the target bound TermSet, and the bound `UNVALIDATED` relationship is recreated with target-local WssIds. |
| `DanglingTermAbsent` | The GUID does not resolve as a live Term, but its exact site-collection cache relationship remains readable. | Keep the GUID absent and recreate its hidden-list relationship with target-local WssIds. |
| `Conflict` / `Unknown` | Live resolution, field binding, WssId, hidden-list identity, label/path, or `TaxCatchAll` evidence is missing or contradictory. | Block. No best-effort repair path is selected. |

Mixed states in one multi-value field are planned independently. A selected field is writable only when every non-empty value has one exact executable action. An unselected field retains one evidence-only action per captured value.

## Captured evidence

For each taxonomy field, the source snapshot records:

- field GUID and internal name;
- source term-store GUID and bound TermSet GUID when readable;
- companion text-field GUID and open/closed setting when readable;
- every value's exact label, Term GUID, and source WssId;
- the WssId row from the site collection's `TaxonomyHiddenList`;
- relevant `TaxCatchAll` hidden-list identity;
- all returned language-specific `Term{LCID}` and `Path{LCID}` values;
- `Title`, `CatchAllData`, and `CatchAllDataLabel`;
- live TermSet, label, path, and tagging evidence when the Term resolves;
- relationship state, capture time, diagnostics, field-value-set digest, and per-relationship evidence digest.

`taxonomyValueSetSha256` binds the field definition, binding, and complete ordered value set. Each `evidenceSha256` binds one relationship proof to that field digest and its hidden/live evidence. Recomputing only the outer snapshot digest cannot conceal a changed taxonomy proof because package validation re-derives both inner digests.

If field binding or live/hidden-list evidence cannot be read, the snapshot stores an incomplete binding identity, marks the relationship `Conflict`, records diagnostics, and seals that observed absence as evidence. Export remains possible. A selected field cannot execute from that evidence; an unselected field can retain it as `RetainEvidenceOnly`.

## Planning and target admission

`TaxonomyTargetMapping` maps source store/TermSet identities to reviewed target store/TermSet identities. It does not authorize Term creation or label-based substitution.

For a field selected for replay, planning requires:

1. one exact source relationship proof;
2. exactly one mapping for the source bound TermSet;
3. for `LiveOutsideBoundTermSet`, exactly one additional mapping for the source live TermSet;
4. a target page field whose exact field GUID, companion text-field GUID, open setting, store, and bound TermSet are sealed into the action;
5. the required target live/absent Term state and exact GUID, label, path, and tagging availability;
6. no contradictory target hidden-list identity.

Import repeats the target-state analysis before mutation. A target Term appearing, disappearing, moving into the bound set, changing identity, or a field binding changing after approval rejects admission. The sealed plan is not rewritten.

`RetainEvidenceOnly` actions are intentionally excluded from target admission, materialization, and fresh target verification. They describe preserved source evidence, not an approved target transaction. If the field later becomes selected, a new plan must evaluate every relationship against the then-current target.

`PagePlanningOptions.BlockOnManagedMetadata` remains serialized for compatibility but is not an escape hatch. Relationship proof, reviewed mappings, admission, and no-repair rules always apply.

## Execution

Execution never calls a Term or TermSet creation API.

For a valid bound relationship, Import assigns the exact existing Term GUID through the target taxonomy field.

For an invalid relationship, Import:

1. freshly verifies that the target still has the approved invalid live-resolution state;
2. creates or exactly reuses target-local `TaxonomyHiddenList` rows;
3. rewrites only the target store/TermSet portions of a valid compressed `CatchAllData` identity; `UNVALIDATED` remains `UNVALIDATED`;
4. maps every captured `Term{LCID}` and `Path{LCID}` value to the same target LCID field and blocks if the target cannot represent a captured locale;
5. rejects multiple or colliding rows instead of overwriting them;
6. writes the original Term GUID and label with the target-local WssId;
7. allows SharePoint's supported write path to normalize the value WssId to the bound `UNVALIDATED` row only when that row is itself part of the sealed action.

The page field binding is not silently changed by this writer. Schema/layout materialization or target preparation must produce the exact binding reviewed during planning.

## Fresh verification and receipt

Storage verification does not trust the successful write call. A cloned context freshly reads:

- the page taxonomy value and observed target WssId;
- target field store/TermSet binding;
- live-in-bound, live-outside-bound, or globally absent Term state;
- the observed `TaxonomyHiddenList` row;
- the required `TaxCatchAll` WssId for invalid relationships.

Each executed result is emitted as `TaxonomyRelationshipVerificationResult`. Every executable relationship must pass for `taxonomyRelationshipsMatched` and aggregate `freshReadbackPassed` to be true. `RetainEvidenceOnly` is not reported as a target verification success because it made no target claim.

## Forbidden repair behavior

The following are intentionally not migration actions:

- creating a missing source Term at the target;
- creating a same-label replacement with a different GUID;
- moving or copying an outside Term into the bound TermSet;
- rebinding a field merely to make an invalid value valid;
- replacing a dangling value with an ancestor, default, or nearest match;
- dropping an invalid value without an explicit future ingredient policy that records reviewed loss;
- overwriting a colliding hidden-list row.

If a future workflow wants remediation, it must define a different explicit policy and action. Remediation must never be presented as faithful reproduction.
