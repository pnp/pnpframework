# Page migration architecture

`PnP.Framework.Migration.Pages` is the shared kernel for staged migration of SharePoint page artifacts. It contains page-wide facts and mechanics that can be reused by classic Publishing Pages, Wiki Pages, Web Part Pages, and later page families.

It does not define one universal page contract. Each page family owns the aggregate that represents its actual storage model, while composing the shared capabilities that apply to it.

The cross-domain workflow, package boundary, and governed-object lifecycle are documented in the Migration [design documentation](../docs/README.md).

## Ownership model

```text
PnP.Framework.Migration.Pages
    shared page identity, ASPX evidence, CLR runtime, profiles, ingredients, and mechanics
        ^
        |
PnP.Framework.Migration.Pages.Publishing
    publishing-page aggregate, layout, lifecycle, package, and verification
        ^
        |
PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
    thin Enterprise Wiki v1 workflow facade and discovery
```

Dependencies must point inward. Shared `Pages` code must not reference `Publishing` or `EnterpriseWiki`. A page-family layer may compose shared page capabilities, and a profile may compose both the page family and the shared capabilities.

This is composition, not an inheritance hierarchy. A future Wiki Page implementation should not inherit a publishing-page snapshot merely to reuse fields or Web Parts; it should define its own aggregate and include the shared models it actually captures.

## Shared capabilities

The folder and namespace layout records ownership:

| Namespace | Shared responsibility |
| --- | --- |
| `Pages` | `PageIdentity` and page path rules. Identity contains web, file, list-item, content-type, version, size, modified time, and title facts that are not tied to one page family. |
| `Pages.Capture` | Source-only capture options, capture status, file probing, and the before/after source stability fence. |
| `Pages.Markup` | Exact source ASPX artifacts, encoding, and parsed Page-directive attributes. |
| `Pages.Runtime` | CLR-first runtime adapter resolution; Content Type is an explicit fallback only. |
| `Pages.Profiles` | Non-exclusive product-profile signals. |
| `Pages.Cohorts` | Versioned workflow validation-population assessments. |
| `Pages.Ingredients` | Canonical ingredient nodes, dependency edges, semantic actions, dependency-closure evaluation, and aggregate outcome. |
| `Pages.Content` | Digest-sealed text replacement descriptions and deterministic text rewriting. This is storage-agnostic; each family chooses which text-bearing properties use it. |
| `Pages.Fields` | Complete list-item field-definition/value evidence, supported value representations, per-field plan actions, and conservative field writes. |
| `Pages.Lifecycle` | Source lifecycle evidence only: checkout type, file level, moderation status, and timestamps. Target lifecycle interpretation belongs to a page family. |
| `Pages.ClassicWebParts` | Shared classic Web Part export evidence, placement, replay capability, and binding/action planning. |
| `Pages.References` | URL/dependency evidence, per-reference actions, text-mapping construction, and approved payload materialization. |
| `Pages.Security` | Permission inheritance and role-assignment evidence. Replay policy belongs to the consuming plan/profile. |
| `Pages.Planning` | Page-wide planning inputs and operation names that are meaningful across page families. |
| `Pages.Packaging` | Internal canonical digest primitives reusable by family-specific packages. |

The shared kernel deliberately does not own:

- publishing layout identity or `PublishingPageContent`;
- Draft/Published target behavior for a publishing page;
- Enterprise Wiki validation-cohort membership or the workflow's preferred stock layout;
- a single package schema for all possible page families.

## Evidence and policy are separate

A type belongs in the shared kernel when it describes a fact or performs a mechanism that remains valid across page families. A decision belongs at the narrowest layer that has enough context to make it.

Examples:

| Shared evidence/mechanism | Policy owner |
| --- | --- |
| `PageLifecycleSnapshot` records source checkout, level, and moderation evidence. | `PublishingPageLifecyclePolicy` decides whether a publishing target can be Published or must remain Draft. |
| `PageArtifactSnapshot` preserves exact ASPX bytes and the Page directive. | `PageRuntimeResolver` selects an adapter from CLR evidence; workflow policy later decides whether that adapter is supported. |
| `PublishingPageProfileSignalProjector` records Content Type/layout/field traits. | `EnterpriseWikiV1CohortPolicy` independently decides EW-v1 validation membership. |
| `ClassicWebPartSnapshotReader` exports Web Part XML and placement. | `ClassicWebPartReplayCapabilityPolicy` and `ClassicWebPartActionPlanner` assess current Publishing replay capability and dependencies. |
| `PageFieldSnapshotReader` captures every returned Pages-library field. | A profile identifies the fields it currently understands; the plan applies only reviewed, compatible actions. |
| `PageReferenceSnapshotReader` captures authored references and safe payloads. | A target plan decides whether each reference is preserved, rewritten, materialized, delegated, or blocked. |
| `PageSecuritySnapshotReader` captures inheritance and role assignments. | The current Enterprise Wiki planning policy requires inherited permissions and does not replay unique assignments. |

This separation preserves unsupported evidence without pretending that every family can restore it in the same way.

## Family composition

A page family defines an aggregate around its real storage shape. The publishing family currently composes:

- common `PageIdentity`;
- exact common ASPX artifact and CLR runtime resolution;
- non-exclusive profile signals and canonical ingredient graph;
- publishing-specific layout evidence and `PublishingPageContent`;
- common field, classic Web Part, reference, security, and lifecycle evidence;
- common source stability fence and capture policy;
- publishing-specific migration plan, package, report, and target verification contracts.

A future Wiki Page aggregate might instead capture a Wiki content field while reusing identity, fields, classic Web Parts, references, security, and lifecycle evidence. A Web Part Page aggregate may emphasize Web Part zones and have no publishing layout. Neither requires duplicate `PageField*`, `ClassicWebPart*`, `PageReference*`, or `PageSecurity*` models.

## Shared contract rules

Shared evidence types follow these rules:

1. Capture records what SharePoint returned, including unknown or currently unsupported values.
2. Capture status and diagnostics describe fidelity; they do not silently discard a value.
3. A plan produces an explicit action for every non-empty canonical ingredient plus the typed domain actions needed for execution.
4. Import applies only actions understood by the current implementation.
5. Family/profile policy cannot be hidden inside a shared reader.
6. Shared mechanics cannot assume a content type, layout name, library template, or target site template.
7. Family-specific aggregate and package schemas may evolve independently when their storage models differ.
8. A retained ingredient cannot drop a required dependency unless a reviewed transform explicitly releases it.

## Adding another page family

Before adding `Pages.Wiki`, `Pages.WebPart`, or another family:

- define the family's source aggregate around its actual storage properties;
- reuse shared evidence types only where their semantics match;
- keep family lifecycle and content behavior in the family namespace;
- keep template/profile classification and portability rules below the family layer;
- do not add profile switches to shared readers;
- define a versioned family package, target plan, and fresh-readback verifier;
- document which evidence can be restored now and which remains recovery-only;
- add dependency-direction checks or code review checks ensuring shared `Pages` never references the new family.

The current Publishing Page implementation is documented in [Publishing/README.md](Publishing/README.md).
