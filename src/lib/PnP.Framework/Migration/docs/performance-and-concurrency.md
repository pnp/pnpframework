# Performance and concurrency

> Status: Draft
> Implementation status: Target-inspection batching implemented; orchestration policy remains host-owned
> Scope: Capture, planning, admission, execution, and verification performance

## Performance is a measured contract

Migration performance must be optimized without changing captured evidence, selected actions, target identities, blockers, warnings, or verification assertions. A faster result is not equivalent when its canonical snapshot or plan digest changes unexpectedly.

Every performance change should therefore establish both:

1. an operational improvement, measured by wall time, stage time, request count, and tail latency;
2. a semantic regression result, measured by canonical digests and typed decisions over the same frozen inputs.

Elapsed time alone is insufficient because batching can accidentally omit a target property, and caching can turn a fresh admission check into stale evidence.

## Measurement stages

A migration host should record a run identity and measure at least these boundaries:

| Stage | Begins | Ends | Useful measures |
| --- | --- | --- | --- |
| Authentication | Before token, cookie, or request-digest acquisition | A usable connection can issue its first request | duration, cache outcome, refresh count |
| Capture | Before opening the source object graph | The sealed source export is written | duration, source requests, bytes, artifact count |
| Planning | Before target inspection | The sealed migration plan is written | duration, target requests, action/blocker count |
| Admission | Before fresh target preflight | Mutation is authorized or rejected | duration, target requests, disposition changes |
| Execution | Before the first mutation intent | The final mutation receipt is recorded | duration per domain/action, retries, written bytes |
| Verification | Before fresh readback | Storage assertions and receipt are complete | duration, read requests, assertion failures |
| Runtime verification | Before external navigation | Runtime receipt is returned | navigation, render, probe, and comparison durations |

For repeated items, report count, mean, P50, P95, maximum, failure count, maximum observed concurrency, and retry count. Network operations should additionally record HTTP status, response bytes, and a safe operation-shape identifier. Request bodies, field values, cookies, tokens, and authorization headers must not enter performance telemetry.

## Optimize SharePoint round trips first

CSOM latency normally dominates local projection and digest computation. The preferred optimization order is:

1. combine compatible `Load` expressions before one `ExecuteQueryRetry`;
2. pass already-resolved objects, such as the Pages library, into downstream planners;
3. reuse properties that are demonstrably loaded on the current `ClientObject`;
4. batch independent file-existence probes, with a compatibility fallback when SharePoint aborts the batch on one missing file;
5. avoid loading broad collections when the source snapshot proves that the domain is empty;
6. apply bounded item-level concurrency only after per-item request count is minimized.

The implementation must not share one `ClientContext` across concurrent items. Each concurrent item owns its contexts and disposes them when its stage ends.

## Freshness and cache boundaries

The following may be cached for their explicit validity period:

- authentication tokens according to their expiry;
- SharePoint FormDigest values according to the context-info timeout, refreshed before expiry;
- immutable local files addressed by a verified SHA-256 digest;
- parsed frozen input manifests inside one process.

The following must not be reused as execution-time truth:

- planning target probes during import admission;
- pre-mutation objects during final verification;
- ownership or collision observations from another item when the target may have changed;
- a previous run's permission, file-existence, List, field, content-type, or lifecycle observation.

Reusing an already-loaded property inside one target-inspection boundary is not a cross-boundary cache. Import admission and verification still create fresh reads.

## Bounded concurrency policy

Concurrency is an orchestration concern because a library call does not know the caller's tenant limits, target concentration, or retry budget. A host may run independent capture or planning items concurrently when every item has its own contexts and output location.

Recommended starting limits are:

| Stage | Default upper bound | Reason |
| --- | ---: | --- |
| Planning | 8 | Read-only, independent target inspections; validate against tenant telemetry. |
| Capture | 4 | Read-only but can transfer larger artifacts and expand dependency graphs. |
| Combined capture-plan | 4 | Bounded by the heavier of source capture and target planning. |
| Apply | 1 per governed target scope | Avoid concurrent ownership races and conflicting dependency writes. |
| Fresh verification | Same operation, sequential after its writes | Preserves a clear freshness fence and receipt ordering. |

These are host defaults, not a promise that every tenant can sustain them. Reduce the limit when P95 latency grows materially, memory or artifact I/O becomes the bottleneck, or retry signals appear. A later scheduler may additionally cap concurrency per tenant, site collection, or Web while retaining a larger global limit.

## Retry and backoff

PnP Framework retry primitives remain authoritative for transient CSOM failures. Hosts should count retries and distinguish authentication/access failures from transient service pressure.

- HTTP 429 and 503 are retry/backoff signals, not evidence that an ingredient is permanently blocked.
- HTTP 401 and 403 identify an authentication or authorization boundary and require renewed credentials, changed permission, or an explicit blocked result at the workflow level.
- An unexpected exception must retain its evidence and operation context; it must not be converted into a semantic migration blocker merely to finish a batch.

Increasing parallelism while retries are rising is not an optimization. Compare throughput together with P95 request latency and total retry delay.

## Semantic regression gate

Before accepting an optimization, run the same frozen snapshot and target mapping before and after the change and compare:

- every `SnapshotDigest` when capture changed;
- every `PlanDigest` when planning or target inspection changed;
- package state, aggregate outcome, action/disposition, blocker, and warning sets;
- target identities and ownership evidence included in the plan;
- import and verification receipts when execution changed.

For a cohort, sort stable `itemIdentity=canonicalDigest` pairs and hash the combined text. Matching cohort hashes provide a compact guard, but reviewers should still inspect representative simple, complex, and fidelity-sensitive packages.

An expected digest change must identify the evidence or decision that changed. A performance-only pull request should normally preserve all canonical digests.
