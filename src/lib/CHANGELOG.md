# PnP Framework Changelog

*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [1.11.0]

### Added

- Added .NET 7 binaries [jansenbe - Bert Jansen]
- Moved to C# 10 [jansenbe - Bert Jansen]
- Adding support to export ListInstance-FieldRef displayName with multilanguages #775 [magarma - Miguel Angel García Martínez]

### Changed

- Fix fileuniqueid and pageuniqueid export #692 [rjbooden - Ronald Booden]
- Added null check for checking if members array is not null #694 [danielpastoor - Daniel Pastoor]
- Add optional parameter to disable welcome message #700 [ohaak2 - Ole Rühaak]
- Calculated field should be created after other fields #627 [friendbool]
- Schema Implementation: ShowPeoplePickerSuggestionsForGuestUsers [pkbullock]
- Schema Implementation: Audience Targeting Classic/Moder [pkbullock]
- Add geo locations that has been recently added #732 [patrikhellgren - Patrik Hellgren]
- Fix ClearDefaultColumnValues working on large lists #717 [cebud - Martin Dubec]
- Export calculated field formula based on title instead of internal name #721 [madsmai - Mads Maibohm]
- Allowing to create PnPContext from ClientContext with existing IPnPContextFactory #762 [heinrich-ulbricht - Heinrich Ulbricht]
- Add support to exporting renamed Title columns to a PnP template #776 [jackpoz - Giacomo Pozzoni]
- Drop unused Microsoft.Extensions.Logging.Abstractions reference [jansenbe - Bert Jansen]
- Check owners if null before calling AddUnifiedGroupMembers #769 [magarma - Miguel Angel García Martínez]
- Fixes for the EnableModernAudienceTargeting method #773 [patrikhellgren - Patrik Hellgren]
- Fix for creating private channels #778 [patrikhellgren - Patrik Hellgren]
- Fix for infinite loop in Replacing Tokens when property values contain the Regex special char "$" #777 [Autophanous]
- Fix LocalizationToken.GetReplaceValue: fallback to old logic #785 [czullu - Christian Zuellig]
- Fix: Failed to resolve termsetid token #786 [czullu - Christian Zuellig]
- Token Regex for fileuniqueid* matches too broadly #751 #763 [heinrich-ulbricht - Heinrich Ulbricht]

## [1.10.0]

### Added

- Added async counterparts for PnP SDK interop. #639 [s-KaiNet - Sergei Sergeev]
- Added GetSitePropertiesById functionality to TenantExtensions #638 [koenzomers - Koen Zomers]

### Changed

- Exporting/Provisioning _ModerationStatus for folders #654 [magarma - Miguel Angel García Martínez]
- Do not try to tokenize non Guid termsetids as they already are tokenized #659 [czullu - Christian Zuellig]
- Adding token parsing of folder name before setting default field values #664 [eduardpaul - Eduard Paul]
- Add ConfigureAwait(false) to AuthenticationManager #665 [RiccardoGDev]
- Fix for using passed PnPMonitoredScope in async method #666 [patrikhellgren - Patrik Hellgren]
- Change to not re-throw caught exception during provisioning #668 [patrikhellgren - Patrik Hellgren]
- Extend timeouts and add retries to teams provisioning #669 [patrikhellgren - Patrik Hellgren]
- Implement show / hide of site title in header using SetChromeOptions #670 [eduardpaul - Eduard Paul]
- Improve folder creation in EnsureFolderPath method #673 [gautamdsheth - Gautam Sheth]
- Add UserAgent to CSOM based access token retrieval in GetAccessToken #642 [andregouveia8 - André Gouveia]
- Remove AddContentTypeHiddenFieldsToList #679 [magarma - Miguel Angel García Martínez]
- Add owners as members too when creating group connected team site using application permissions #680 [gautamdsheth - Gautam Sheth]
- Fix for empty PnP:Templates in the xml #677 [danielpastoor - Daniel Pastoor]
- Fix for getting the teams photo when using application permissions #678 [danielpastoor - Daniel Pastoor]
- Ensure correct PnPContext is returned from the passed ClientContext #676 [danielpastoor - Daniel Pastoor]
- Fix #1696 - issue with team creation , use graph token if possible #681 [gautamdsheth - Gautam Sheth]
- Avoid retrying when the hostname cannot be found #686 [koenzomers - Koen Zomers]

## [1.9.0]

### Added

- Support export folder default values #583 [magarma - Miguel Angel García Martínez]
- Added RequestTemporaryAccessPass method to UsersUtility #605 [koenzomers - Koen Zomers]
- Transformation: added specific page layout transformation #614 [robi26 - Stephan Steiger]

### Changed

- Feature : remove special chars from the SiteAlias #558 [gautamdsheth - Gautam Sheth]
- Fix: set ContextSettings Type to ClientContextType.PnPCoreSdk #566 [czullu - Christian Zuellig]
- Fix: for adding folders in large list #559 [gautamdsheth - Gautam Sheth]
- Added localization support for child nodes in footer #560 [b00johst - John Frankolin]
- Additional changes for fix #498, improving header provisioning scenario #562 [fzbm - Florian Zink]
- Catching File.NotFound exception to prevent error during export #573 [magarma - Miguel Angel García Martínez]
- Check for null before foreach to avoid nullExceptionIssue #574 [magarma - Miguel Angel García Martínez]
- Fix Clear-PnPDefaultColumnValues not working with taxonomy fields #576 [jackpoz - Giacomo Pozzoni]
- Files tab teams first load after creation #588 [roberAlb - Roberto Ramon]
- Adding beta option to MS Graph Users functions #586 [koenzomers - Koen Zomers]
- Added ignoreDefaultProperties to UsersUtility #590 [koenzomers - Koen Zomers]
- Rewrite of the get user delta #602 [koenzomers - Koen Zomers]
- Fix ProcessDataRow ignore html fields as empty even they are not #604 [czullu - Christian Zuellig]
- Fix caching issues with retrieving taxonomy fields on modernization caching #603 [erosu-hab - Elena Rosu]
- Retrieve the target list by url for handler ObjectListInstanceDataRows #607 [BollietMZK]
- Fix exception thrown when a Taxonomy Term owner was not found #628 [jackpoz - Giacomo Pozzoni]

## [1.8.0]

### Added

- Added Sensitivity Labels to Unified Group Creation #536 [NikoMix - Niko]
- Added support to provisioning site collections using the classic site collection creation APIs from within a tenant template [jansenbe - Bert Jansen]

### Changed

- Fixing GetGroupMembers and GetGroupOwners #490 [koenzomers - Koen Zomers]
- Fix for false positive exception logging in GetPrincipalUniqueRoleAssignments #495 [patrikhellgren - Patrik Hellgren]
- Fix for missing token for newly created list content type #496 [patrikhellgren - Patrik Hellgren]
- Fix for issue #500, unable to ensure folder with path characters #501 [jimmywim - Jim Love]
- Use redirectUrl if specified #503 [wobba - Mikael Svenson]
- Support default column lookup #504 [wobba - Mikael Svenson]
- Fix for issue with empty description in UserCustomActions #505 [czullu - Christian Zuellig]
- Fix UrlEncode/Decode Issues with special char #506 [czullu - Christian Zuellig]
- Fix for IsFeatureActiveInternal when Feature does not exist for the Template #507 [czullu - Christian Zuellig]
- Add TEAMCHANNEL#1 to BaseTemplates #508 [czullu - Christian Zuellig]
- Implemented token parsing of team channel name and description #510 [patrikhellgren - Patrik Hellgren]
- Adding top instruction for ListUsers to avoid throttling #513 [koenzomers - Koen Zomers]
- CreateFolderInList Handle Error "To update this folder, go to the channel in Microsoft Teams" #470 [czullu - Christian Zuellig]
- Switched to using version 1.9 of image webpart, needed to ensure images are correctly sized on target pages [jansenbe - Bert Jansen]
- Retry on SocketExcetion in SendAsync #528 [patrikhellgren - Patrik Hellgren]
- Fixing issue when trying to export datarow for list instances #531 [magarma - Miguel Angel García Martínez]
- Fix issues exporting channels and tabs with TeamsAppId and Description #532 [magarma - Miguel Angel García Martínez]
- Changed the property we get when exporting Teams apps #533 [magarma - Miguel Angel García Martínez]
- Fix for Null Exception when trying to set a term value in a file, listItem, and the term label contains a comma #541 [magarma - Miguel Angel García Martínez]
- Missing HeaderLayoutType implemented #498 [jansenbe - Bert Jansen]
- Fix equality comparisons in Provisioning model objects #410 [orty - Serge ARADJ]
- When provisioning pages allow creation of web parts which are not returned as possible web part to add [jansenbe - Bert Jansen]
- Fix #547 - issue with trace log #548 [gautamdsheth - Gautam Sheth]
- Fix issue when folderpath doesn't start by '/' #542 [magarma - Miguel Angel García Martínez]
- Added missing field Type "Geolocation" case in UpdateListItem() #546 [PedroMordeP - Pedro Monte]
- Setting Teams channel as private according to membershipType #549 [magarma - Miguel Angel García Martínez]
- Fix for Culture is not supported exception #554 [patrikhellgren - Patrik Hellgren]
- Added progress logging for provisioning of extensibility handlers #555 [patrikhellgren - Patrik Hellgren]
- Fix for unauthorized exception in EnsureFolder #556 [patrikhellgren - Patrik Hellgren]

## [1.7.0]

### Added


### Changed

- Adding operating system in order to fix issue #737 #440 [koenzomers - Koen Zomers]
- Fixed issue preventing term store data being cached during page transformation [jansenbe - Bert Jansen]
- Updates done to make PnP Framework work without having a dependency on the PnP Core SDK internals [jansenbe - Bert Jansen]
- Updated build-debug script and project file that enables building and testing PnP Framework with a locally build PnP Core SDK assembly [jansenbe - Bert Jansen]
- GetWebSearchCenterUrl makes us loose the pending changes in CSOM Context #454 [czullu - Christian Zuellig]
- Update base templates for the provisioning engine [jansenbe - Bert Jansen]
- Including the requested SelectProperties in the output of ListUsers #460 [koenzomers - Koen Zomers]
- Fix PnPPS #1024 - issue with folder name having special chars #461 [gautamdsheth - Gautam Sheth]
- Fix regression - adding users/groups as site collection admins #462 [gautamdsheth - Gautam Sheth]
- PnP Framework assembly is now strong named #458 #468 [jansenbe - Bert Jansen]
- Fixing issue with SetDefaultColumnValuesImplementation not working for folders with special characters in them #471 [koenzomers - Koen Zomers]
- Feature - additional changes for files/folders with special chars #476 [gautamdsheth - Gautam Sheth]
- Fix: Added additional check to avoid exception when footer not enabled on a site #479 [koenzomers - Koen Zomers]
- Fix: UsersUtility selectProperties bugfix #480 [koenzomers - Koen Zomers]

## [1.6.0]

### Added

- Add support to ClearExistingItems attribute of <pnp:Members> for custom groups #386 [jackpoz - Giacomo Pozzoni]
- Added configuration settings for SharePoint tab entity inside Teams #385 [roberAlb]
- Added new provisioning token for using sequence site tokens in global template context #384 [patrikhellgren - Patrik Hellgren]

### Changed

- Fix TokenParser always returns original associated-group-names #422 [czullu - Christian Zuellig]
- Better fix for issue 390. Now it waits 5 secs between each retry. #418 [luismanez - Luis Manez]
- Fix pnpcontext cache did not work for multitenant #417 [czullu - Christian Zuellig]
- Fix for issue setting ClientSidePage taxonomy field values from parameters empty #407 [magarma - Miguel Angel García Martínez]
- ProcessFields was using wrong Token-Parser #403 [czullu - Christian Zuellig]
- Issue 390 fixed in CreateOrUpdateTeamFromGroupInternal. #391 [luismanez - Luis Manez]
- Fix for content type UpdateChildren not being used #387 [patrikhellgren - Patrik Hellgren]
- Fix so that teams apps are added before channels #382 [patrikhellgren - Patrik Hellgren]
- Added JsonConvertor decorator to UpdateBehavior #380 [thechriskent - Chris Kent]
- Fix pnpcoresdk instance GetClientContext init ContextSettings #370 [czullu - Christian Zuellig]

## [1.5.0]

### Added

- Support for creating a CSOM ClientContext from a PnP Core SDK context [jansenbe - Bert Jansen]
- Minimal support for TEAMCHANNEL Template #268 [czullu - Christian Zuellig]
- PnPSDK Mocking + few more cases isolated #262 [mgwojciech - Marcin Wojciechowski]
- Added implementation of SPWebRequestExecutor that utilizes HttpClient #261 [patrikhellgren - Patrik Hellgren]

### Changed

- Upgrade Microsoft.Identity.Client.Extensions.Msal, Microsoft.Graph, Microsoft.Identity.Client and System.IdentityModel.Tokens.Jwt references #367 [gautamdsheth - Gautam Sheth]
- Fix Access Denied error when processing <pnp:SiteSettings> object #360 [jackpoz - Giacomo Pozzoni]
- Fix: Restore List.EnableFolderCreation value after creating folders #358 [jackpoz - Giacomo Pozzoni]
- Fix: Only specify LCID when needed during team site creation #354 [gautamdsheth - Gautam Sheth]
- Fix for handling folder default values that are empty in provisioning template #346 [patrikhellgren - Patrik Hellgren]
- Fix for valid characters and lower case in site alias #345. Improves #326. [patrikhellgren - Patrik Hellgren]
- Added HostUrl token for Hubsite url property (fixes #338) #343 [gautamdsheth - Gautam Sheth]
- Fix for issue with provisioning default values for root folder #342 [patrikhellgren - Patrik Hellgren]
- Fix #275 - for tenantId missing issue #328/#350 [gautamdsheth - Gautam Sheth]
- Improve alias naming for modern sites. #326 [gautamdsheth - Gautam Sheth]
- LCID validation while creating site collections #325 [gautamdsheth - Gautam Sheth]
- Enhance pnpcoresdk - preseve PnPContext when init with GetClientContext for GetContext #321 [czullu - Christian Zuellig]
- Fixed a bug that updateChildren value was ignored in AddContentTypeToList #318 [antonsmislevics - Antons Mislevics]
- Fix for token parsing content type names in console and logs #316 [patrikhellgren - Patrik Hellgren]
- Fix an issue with possible access of null serverobject in FieldUtilities.FixLookupField #310 [NickSevens - Nick Sevens]
- Fixes tokenization of SharePoint Syntex models when being exported [jansenbe - Bert Jansen]
- Fixed DeviceLogin in AuthenticationManager for single tenant scenario #308 [PaoloPia - Paolo Pialorsi]
- Improved handling of socket exceptions in the ExecuteQueryRetryImplementation #301 [czullu - Christian Zuellig]
- Match JsonPropertyName values in ExtractConfiguration with https://aka.ms/sppnp-extract-configuration-schema #300 [jackpoz - Giacomo Pozzoni]
- Fix tokenize url in new document template #299 [czullu - Christian Zuellig]
- Fix list id token issue when title null #298 [czullu - Christian Zuellig]
- Fix conversion between PnPCore enum values and PnP.Framework enums #297 [jackpoz - Giacomo Pozzoni]
- Fix IEnumerable.Union() result being discarded in ContentByQuerySearchTransformator #296 [jackpoz - Giacomo Pozzoni]
- Improve exception messages when acsTokenGenerator is null in AuthenticationManager #295 [jackpoz - Giacomo Pozzoni]
- Fix hanging request and big list issue with CreateDocumentSet #290 [YannickRe - Yannick Reekmans]
- Fix #271: Added PDL support for creating sites via PnP PowerShell #278 [gautamdsheth - Gautam Sheth]
- Fix issue if SP Group has description with more then 512 char #270 [czullu - Christian Zuellig]
- Fix: Always create private channels with isFavoriteByDefault false #260 [patrikhellgren - Patrik Hellgren]
- Fix: Implemented retry logic for getting and creating teams tabs #259 [patrikhellgren - Patrik Hellgren]
- Fix to add groups as owners of site collections #258 [gautamdsheth - Gautam Sheth]
- Fix url token root site #256 [czullu - Christian Zuellig]

## [1.4.0]

### Added

- Added REST mocking scenario + isolation for two test suites #221 [mgwojciech - Marcin Wojciechowski]
- Added Get/Add EventReceiver method to SiteExtensions #166 [bhishma - Bhishma Bhandari]

### Changed

- Fix to always copy WebRequestExecutorFactory when cloning #255 [patrikhellgren - Patrik Hellgren]
- Align search navigation node deletion with other structural navigation node deletion [jansenbe - Bert Jansen]
- Added NeutralResourcesLanguage assembly attribute [jansenbe - Bert Jansen]
- Deserialze of Template XML does not load ClientSideComponentId on List-UserCustomAction #230 [czullu - Christian Zuellig]
- Fix Site Design creation/update via provisioning engine - WebTemplate not being set #229 [michael-jensen Mike Jensen]
- Fix AudienceUriValidationFailedException exceptions when using AppCatalogScope.Tenant #228 [jackpoz - Giacomo Pozzoni]
- Allow to specify the Sensitivity Label Id when creating a new site collection #226 [jackpoz - Giacomo Pozzoni]
- Fix for ListItem.GetFieldValueAs not working for other types than string #223 [patrikhellgren - Patrik Hellgren]
- Documentation updates #215 [LeonArmston - Leon Armston]

## [1.3.0]

### Added

### Changed

- Adding functionality to deal with all types of Azure Active Directory groups #175 [koenzomers - Koen Zomers]
- Adding VS Code build profile #176 [koenzomers - Koen Zomers]
- Explicitely set UseCookies to false when creating a HttpClient since .NET Framework requires this (versus .NET Core/.NET 5 that work without this setting) [jansenbe - Bert Jansen]
- Only assume missing levels in content type hiarchy when the list was created from an OOB list template [jansenbe - Bert Jansen]
- Enable applying of a theme at web level [jansenbe - Bert Jansen]
- Fix: Changed the way we check the teams tab app id #189 [patrikhellgren - Patrik Hellgren]
- Fix for updating team if a tab had Remove = True #190 [patrikhellgren - Patrik Hellgren]
- Allow to have Custom UserAgent when use PnPHttpClient.Instance.GetHttpClient #195 [czullu - Christian Zuellig]
- Hidden webparts can now be skipped during transformation #194 [jansenbe - Bert Jansen]
- Improved Tenant Id fetch method #198 [gautamdsheth - Gautam Sheth]
- Fix if file is template in /libname/Forms folder and therefore has no ListItemId #206 [czullu - Christian Zuellig]
- Allow for setting the http client timeout via SharePointPnPHttpTimeout as environment variable or app setting [jansenbe - Bert Jansen]

## [1.2.0]

### Added

- Added CreateWith* methods to AuthenticationManager to help creation of new AuthenticationManager objects
- Refactored AuthenticationManager to use the library wide single instance of the HttpClient.
- Support for using the ICustomWebUi to enable interactive auth in PnP PowerShell [erwinvanhunen - Erwin van Hunen]
- Support for offline CSOM testing! #100 [mgwojciech - Marcin Wojciechowski]
- Support for on-premises context creation and cloning, internal only as it's needed to support page transformation from on-prem [jansenbe - Bert Jansen]
- Allow use of tokens for TopicHeader and AlternativeText in Pages #137 [magarma - Miguel Angel Garc�a Mart�nez]

### Changed

- Fixed page transformation caching manager to handle dictionary serialization with non string value as keys. Fixes #110 [jansenbe - Bert Jansen]
- Fix http exception unwrapping, refactor http retry mechanism #107 [sebastianmattar - Sebastian Mattar]
- Fixed page transformation caching manager to handle another dictionary serialization with non string value as keys issue. Fixes #136 [jansenbe - Bert Jansen] 
- Refactor ACS token creation #119 [sebastianmattar - Sebastian Mattar]
- Fix wrong ReadOnly setting on ContentType LinkedFields #128 [czullu - Christian Zuellig]
- Fix for #117: failed to set content type on client side page #135 [czullu - Christian Zuellig]
- Fix bug when trying to add existing member or owner to a unified group. #139 [magarma - Miguel Angel Garc�a Mart�nez]
- Fix: Updating WebTemplateExtensionId value in payload dictionary. #143 [magarma - Miguel Angel Garc�a Mart�nez]
- Fix some warnings #147 [jackpoz - Giacomo Pozzoni]
- Fix Escaped whiteSpace break JSON in NewDocumentTemplates #152 [czullu - Christian Zuellig]
- Fix because Modern List Creation creates CTType with 0x0100ParentOne00Id but ParentOne does not exist #153 [czullu - Christian Zuellig]
- Updated UsersUtility to retrieve all users when requested #157 [koenzomers - Koen Zomers]
- Replace Thread.Sleep() with Task.Delay() in async methods #165 [jackpoz - Giacomo Pozzoni]
- Added additional Graph User properties #160 [koenzomers - Koen Zomers]

## [1.1.0]

### Added

### Changed

- New release due to change in PnP Core SDK that should get included

## [1.0.0]

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
- Fix - sensitivity label + preferred data location missing changes #50 [gautamdsheth - Gautam Sheth]
- Fix for unhandled null reference exception #58 [patrikhellgren - Patrik Hellgren]
- Feature - update beta endpoints to v1.0 for UnifiedGroups methods #62 [gautamdsheth - Gautam Sheth]
- Feature - added additional props for Team Site No Group #63 [gautamdsheth - Gautam Sheth]
- Fix - Parse tokens in the SearchCenterUrl #72 [YannickRe - Yannick Reekmans]
- Cert loading improvements #68 [mbakhoff - M�rt]
- Fix - Keep the existing stack information on rethrowing the exception. #83 [YannickRe - Yannick Reekmans]
- Fix a null reference exception in cases where ClientContextSettings are null. #82 [YannickRe - Yannick Reekmans]
- Fix - app-only issue for teamifying group sites #78 [gautamdsheth - Gautam Sheth]
- Feature - minor improvements related to Graph #77 [gautamdsheth - Gautam Sheth]
- Feature - added support for chunked uploading of files #59 [gautamdsheth - Gautam Sheth]
- Feature - improved best match implementation of content type id #61 [jensotto - Jens Otto Hatlevold]
- Added token parsing of team displayname in log message #96 [patrikhellgren - Patrik Hellgren]
- Added check for existing team before checking archived status #95 [patrikhellgren - Patrik Hellgren]
