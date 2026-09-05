# Taxonomy relationship fidelity

> Status: Draft
> Implementation status: Publishing Page relationship replay plus taxonomy asset capture, planning, target inspection, approval, journaled materialization, and fresh verification are implemented; the 10% live plan remains unapplied pending review
> Applies to: page-field taxonomy capture, taxonomy asset closure, planning, import, and fresh verification

## Decision

Migration preserves the relationship observed at the source. It does not repair taxonomy.

An invalid source relationship is therefore not treated as a request to invent a Term that is absent at the source, move a live outside-bound Term into the field's bound TermSet, substitute a same-label Term, or silently drop the value. If the owning field is selected for replay and the exact relationship can be proven and reproduced, the plan contains an explicit relationship action. If a selected relationship cannot yet be reproduced, that ingredient remains a nonterminal planning gap unless retained literal wire evidence proves HTTP `401` or `403` for that branch.

This rule applies independently to every taxonomy value. The page field is the consumer; each value relationship is a required taxonomy ingredient in the canonical dependency graph.

Capture scope remains wider than restore scope. When the current workflow does not select a taxonomy field for replay, every captured value receives `RetainEvidenceOnly`. That action preserves the sealed relationship proof for a later workflow, makes no target capability claim, performs no target taxonomy probe, and projects to the ingredient-level `Delegate` disposition. Source conflict evidence in an unselected field therefore remains recoverable without blocking unrelated page migration.

## Source relationship states

| State | Source fact | Selected-field target interpretation |
| --- | --- | --- |
| `LiveInBoundTermSet` | The Term GUID resolves live inside the field's bound TermSet. | Reuse the same GUID in the explicitly mapped target bound TermSet. Before page admission, the exact captured GUID, label, path, and ancestry must either already exist or have been created by a separately approved exact asset plan. |
| `LiveOutsideBoundTermSet` | The GUID is absent from the bound TermSet but resolves as a live global Term in another TermSet. The page's `TaxCatchAll` retains an `UNVALIDATED` row for the bound TermSet. | Reproduce the same outside-bound relationship. Both source TermSets require reviewed mappings. The exact live Term must remain in the mapped outside TermSet, whether reused or prepared by a separately approved exact asset plan; the bound `UNVALIDATED` relationship is recreated with target-local WssIds. |
| `DanglingTermAbsent` | The GUID does not resolve as a live Term, but its exact site-collection cache relationship remains readable. | Keep the GUID absent and recreate its hidden-list relationship with target-local WssIds. |
| `Conflict` / `Unknown` | Live resolution, field binding, WssId, hidden-list identity, label/path, or `TaxCatchAll` evidence is missing or contradictory. | Defer and retain the exact contradictory/incomplete evidence. No best-effort repair path is selected. Only literal wire HTTP 401/403 evidence may authorization-block the affected branch. |

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

## Taxonomy asset closure

Page bundles retain the complete captured field/value evidence. The cohort taxonomy asset snapshot is intentionally narrower: it contains only affected TermSets plus the exact live Terms and ancestor chain required by selected consumers. The affected TermSet boundary may be discovered from nonterminal `MitigationPending` ingredients, but the consumer/value scan must still cover the complete selected cohort. A page that has no taxonomy mitigation row can carry an additional required value for an affected TermSet.

Asset identity is CLR-typed and GUID-first. Plans retain the source tenant, TermStore, deterministic target ownership TermGroup, TermSet, Term, parent, label/path, language, tagging/open settings, evidence digest, consumers, and Repro4-compatible `pnp_reserved_term_original_identifier` provenance. For every captured live Term they also retain SharePoint's raw `IsReused`, `IsSourceTerm`, reported `SourceTerm`, complete TermSet-membership list, and pin-source TermSet evidence. They never infer identity from a label alone.

Term identity equality is not sufficient for relationship fidelity. An existing same-GUID target Term is equivalent only when its native/reused state, source-Term relationship, translated TermSet memberships, and pin relationship match the source evidence. The source owning TermSet ID is translated through the reviewed target mapping; other captured relationship identities remain exact. A same-GUID Term that would become reused or multi-member in order to fit a second target TermSet is a relationship change, not an exact reproduction.

Read-only target inspection assigns one explicit disposition per asset:

- `CreateMissing`: recreate an exact source-live asset with its preserved source GUID and provenance in a migration-owned target boundary;
- `ReuseOwned`: reuse the single exact provenance-owned asset;
- `ReconcileOwnedPlanDrift`: continue RCA and reconcile an owned asset whose reviewed shape drifted;
- `ReviewExternalReuse`: require explicit approval before mapping to an equivalent asset that lacks migration provenance;
- `CreateMissingAfterExternalApproval`: create an exact source-live child only after the containing external TermSet mapping and mutation are approved;
- `ResolveCollision` and `RetryRequired`: continue evidence collection and mitigation; neither is a terminal authorization result;
- `AuthorizationBlocked`: stop only the affected branch, and only with retained literal wire HTTP `401` or `403` evidence.

Missing topology or taxonomy assets are creation actions, not mapping failures. Asset planning must not mutate the target. Every target ownership TermGroup is an explicit ingredient with its own probe, approval action, mutation journal entry, receipt, and verification result. A TermSet approval therefore cannot hide an implicit TermGroup creation. Materialization is admitted from the exact reviewed plan and approval digests, writes an intent before each mutation, supports exact owned recovery after an interrupted attempt, and requires aggregate fresh readback before page relationship admission can consume the prepared identities. Fresh inspection and final verification reject drift in native/reused state, reported SourceTerm, exact translated membership list, or pin state even when GUID, label, and path still match.

## Planning and target admission

`TaxonomyTargetMapping` maps source store/TermSet identities to reviewed target store/TermSet identities. The mapping alone does not authorize Term creation, external mutation, repair, or label-based substitution. Creation authority can come only from a separate digest-sealed taxonomy asset action.

For a field selected for replay, planning requires:

1. one exact source relationship proof;
2. exactly one mapping for the source bound TermSet;
3. for `LiveOutsideBoundTermSet`, exactly one additional mapping for the source live TermSet;
4. a target page field whose exact field GUID, companion text-field GUID, open setting, store, and bound TermSet are sealed into the action;
5. the required target live/absent Term state and exact GUID, label, path, and tagging availability, including any separately executed and freshly verified asset action;
6. no contradictory target hidden-list identity.

Import repeats the target-state analysis before mutation. A target Term appearing, disappearing, moving into the bound set, changing identity, or a field binding changing after approval rejects admission. The sealed plan is not rewritten.

`RetainEvidenceOnly` actions are intentionally excluded from target admission, materialization, and fresh target verification. They describe preserved source evidence, not an approved target transaction. If the field later becomes selected, a new plan must evaluate every relationship against the then-current target.

`PagePlanningOptions.BlockOnManagedMetadata` remains serialized for compatibility but is not an escape hatch. Relationship proof, reviewed mappings, admission, and no-repair rules always apply.

## Execution

Publishing Page relationship execution never calls a TermGroup, TermSet, or Term creation API. It consumes identities that already passed the separate taxonomy asset admission and verification boundary.

The asset materializer creates or reconciles only an explicitly approved action. It re-probes immediately before mutation, orders `TermGroup -> TermSet -> Term -> child Term`, preserves source GUID and ancestry, records write-ahead mutation intents and receipts, and freshly inspects the complete approved closure afterward. `ReviewExternalReuse` performs no mutation. `CreateMissingAfterExternalApproval` requires both approval of the external TermSet mapping and a separate explicit external-mutation flag on that Term action.

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

- creating a Term whose source relationship was captured as `DanglingTermAbsent`;
- creating a same-label replacement with a different GUID;
- moving or copying an outside Term into the bound TermSet;
- rebinding a field merely to make an invalid value valid;
- replacing a dangling value with an ancestor, default, or nearest match;
- dropping an invalid value without an explicit future ingredient policy that records reviewed loss;
- overwriting a colliding hidden-list row.

Creating an exact source-live Term, with the same GUID and ancestor path, in its separately reviewed mapped TermSet is asset reproduction rather than repair. It becomes repair if the Term was absent at the source, receives a different identity, is moved into the bound set to make an invalid relationship valid, or otherwise changes the captured relationship.

If a future workflow wants remediation, it must define a different explicit policy and action. Remediation must never be presented as faithful reproduction.
