# PnP Framework Changelog

*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [Unreleased]

### Added

- Added Azure OnBehalfOf token acquiring #17 [titns - TTs]
- Added support to extract all teams #27 [gautamdsheth - Gautam Sheth]
- Added SiteAlias to Site Collection creation #29 [patrikhellgren - Patrik Hellgren]
- Added support to document a provisioning template in MD format #39 [kevmcdonk - Kevin McDonnell]

### Changed

- When creating Team through a tenant template using a delegate token the owner of the group that is being created will be set ot the user identified by the token. If an app-only token is provided and no owners are specified in the template and exception will be thrown. [erwinvanhunen - Erwin van Hunen]
- Fix for issue where FooterLink in a sitetemplate get provisioned in reverse order [erwinvanhunen - Erwin van Hunen]
- Removed obsolete Responsive UI methods [erwinvanhunen - Erwin van Hunen]
- Fix for instantiation of extensibility handlers #5 [patrikhellgren - Patrik Hellgren]
- Fix for token parsing when provisioning folders #6 [patrikhellgren - Patrik Hellgren]
- Fix some warnings related to XML comments #9 [jackpoz - Giacomo Pozzoni]
- Fix some warnings #10 [jackpoz - Giacomo Pozzoni]
- fix url encoding issue when writing href values to client_LocationBasedDefaults.html as part of SetDefaultColumnValuesImplementation() #11 [Jaap Vossers - jvossers]
- Fix issue with handling of terms with comma and provided GUID #14 [reusto]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2797 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2759 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2763 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2746 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2790 [gautamdsheth - Gautam Sheth]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2784 [gautamdsheth - Gautam Sheth]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2775 [gautamdsheth - Gautam Sheth]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2762 [gautamdsheth - Gautam Sheth]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2760 [gautamdsheth - Gautam Sheth]
- Add some missing XML comments #20 [jackpoz - Giacomo Pozzoni]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2802 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2801 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2803 [jansenbe - Bert Jansen]
- Fixed NullReferenceException when adding Webhook Subscription #24 [martinewald]
- Replaced graph beta call with already existing private method with v1 graph call #28 [patrikhellgren - Patrik Hellgren]
- Enabled metadata copy of taxonomy and user fields during in-place page modernization [jansenbe - Bert Jansen]
- Fix - Honour Overwrite attribute on Package in Tenant template #33 [YannickRe - Yannick Reekmans]
- Feature - replaced GetFileByServerRelativeUrl to GetFileByServerRelativePath method #31 [gautamdsheth - Gautam Sheth]
- Improvements - removed some extra checks + fix obsolete Telemetry API call #32 [gautamdsheth - Gautam Sheth]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2583 [jansenbe - Bert Jansen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2785 [erwinvanhunen - Erwin van Hunen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2756 [erwinvanhunen - Erwin van Hunen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2819 [erwinvanhunen - Erwin van Hunen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2755 [erwinvanhunen - Erwin van Hunen]
- Ported https://github.com/pnp/PnP-Sites-Core/pull/2745 [erwinvanhunen - Erwin van Hunen]
- Fix ACS token generation #41 [sebastianmattar - Sebastian Mattar]
- Fix some warnings #45 [jackpoz - Giacomo Pozzoni]
- Feature - Additional underlying changes to API #49 [gautamdsheth - Gautam Sheth]

