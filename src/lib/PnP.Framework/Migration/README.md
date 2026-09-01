# Migration

The `PnP.Framework.Migration` namespace contains staged, evidence-preserving migration workflows for SharePoint artifacts.

Migration is intentionally separate from the existing provisioning and modernization engines:

| Area | Primary purpose |
| --- | --- |
| Provisioning | Describe or apply a desired SharePoint configuration. |
| Modernization | Transform supported source experiences into modern SharePoint experiences. |
| Migration | Capture source evidence, make target-specific decisions, apply only an approved plan, and verify the result from a fresh readback. |

The migration APIs are being introduced incrementally. A namespace or contract is not considered stable merely because it is public; schema versions and release notes define compatibility once a feature is released.

## Design rules

Every migration area should preserve the following boundaries:

1. **Capture is source-only.** Export must not require a target connection and must not make target decisions.
2. **Planning is target-specific and read-only.** A planner inspects the target, records every decision, and produces blockers and warnings without changing the target.
3. **Import executes a sealed plan.** The importer validates both the source snapshot digest and the approved plan digest before writing.
4. **Verification uses a fresh readback.** Success is based on persisted target state, not only on successful CSOM requests.
5. **Unknown evidence is retained.** Capture should preserve information even when the current importer cannot restore it. Planning decides what is understood and safe to apply.
6. **Unsafe ambiguity becomes a blocker or a conservative result.** A migration profile must not silently guess cross-site identities, target bindings, or lifecycle state.
7. **Profiles compose reusable object domains.** Site-template-specific behavior belongs in a profile namespace; field, Web Part, reference, security, lifecycle, package, and verification behavior belongs in reusable namespaces.
8. **Use existing PnP primitives.** Migration code should compose established PnP Framework operations instead of duplicating CSOM retry, file, folder, page, URL, or Web Part plumbing.

## Current areas

| Namespace | Scope |
| --- | --- |
| [`PnP.Framework.Migration.PublishingPages`](PublishingPages/README.md) | Classic publishing-page capture, target planning, import, and verification. |
| `PnP.Framework.Migration.PublishingPages.EnterpriseWiki` | Enterprise Wiki classification and target policy built on the publishing-page domains. |

Future artifact families should be peers of `PublishingPages`, not additions to the Enterprise Wiki profile.

## Namespace and folder ownership

Folders represent an object domain or a public workflow boundary, not a generic implementation role. Avoid broad folders such as `Helpers`, `Managers`, `Readers`, or `Writers` when the type naturally belongs to a domain such as `Fields`, `WebParts`, or `Lifecycle`.

A profile may coordinate several domains, but it should not absorb their models or reusable policies. Conversely, generic code must not contain profile-specific content type IDs, layout names, or target assumptions.

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
