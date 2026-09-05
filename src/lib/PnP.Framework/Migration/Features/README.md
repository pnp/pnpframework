# Platform feature migration

This namespace models SharePoint platform features that are conditional prerequisites for captured object relationships. A platform-owned content type is not assumed to exist merely because its ID is known: its provisioning feature must be represented, probed, approved, activated when required, and freshly verified.

The current catalog maps Asset Library, Document Sets, and Video and Rich Media content-type parents to site-scoped feature requirements. Each `PlatformFeatureMaterializationPlan` records:

- feature ID, name, and scope;
- dependency order and required feature IDs;
- captured content-type relationships that require it;
- runtime content-type IDs the feature promises;
- target Site Collection URL and planning-time probe;
- the reviewed ensure-active disposition and reason.

List execution merges identical requirements inside one source/target Site Collection, activates them in dependency order, and verifies the promised runtime content types before site-content-type or List membership actions run. Activation is idempotent: an already active feature is recorded as already satisfied.

Feature capability is evaluated independently from its consumers. A blocked List collision does not make an activatable feature incompatible; the ingredient dependency graph keeps the List gated. Conversely, an active feature whose promised runtime content types are absent is a contract mismatch and cannot be treated as satisfied.

Authorization classification belongs to the surrounding workflow. Only a retained literal HTTP 401/403 exchange may terminate a transaction as authorization-blocked. `E_ACCESSDENIED` inside an HTTP 200 CSOM payload remains an RCA/mitigation result.
