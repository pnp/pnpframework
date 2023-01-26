﻿using System;

namespace PnP.Framework
{
    /// <summary>
    /// Constants.
    /// </summary>
    /// Recommendation: Constants should follow C# style guidelines and be Pascal Case
    public static class Constants
    {
        [Obsolete("Use Constants.FeatureId_Site_AppSideLoading instead.")]
        public static readonly Guid APPSIDELOADINGFEATUREID = new Guid("AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D");
        [Obsolete("Use Constants.FeatureId_Web_MinimalDownloadStrategy instead.")]
        public static readonly Guid MINIMALDOWNLOADSTRATEGYFEATUREID = new Guid("87294c72-f260-42f3-a41b-981a2ffce37a");

        // PublishingWeb SharePoint Server Publishing - Web
        public static readonly Guid FeatureId_Web_Publishing = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
        // PublishingSite SharePoint Server Publishing Infrastructure - Site
        public static readonly Guid FeatureId_Site_Publishing = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
        // MDSFeature
        public static readonly Guid FeatureId_Web_MinimalDownloadStrategy = new Guid("87294c72-f260-42f3-a41b-981a2ffce37a");
        // EnableAppSideLoading
        public static readonly Guid FeatureId_Site_AppSideLoading = new Guid("AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D");

        internal const string SITEFOOTER_NODEKEY = "13b7c916-4fea-4bb2-8994-5cf274aeb530";
        internal const string SITEFOOTER_TITLENODEKEY = "7376cd83-67ac-4753-b156-6a7b3fa0fc1f";
        internal const string SITEFOOTER_LOGONODEKEY = "2e456c2e-3ded-4a6c-a9ea-f7ac4c1b5100";
        internal const string SITEFOOTER_MENUNODEKEY = "3a94b35f-030b-468e-80e3-b75ee84ae0ad";

        internal const string LOGGING_SOURCE = "PnP.Framework";
        internal const string LOGGING_SOURCE_FRAMEWORK_PROVISIONING = "PnP Provisioning";

        internal const string FIELD_XML_FORMAT = @"<Field Type=""{0}"" Name=""{1}"" DisplayName=""{2}"" ID=""{3}"" Group=""{4}"" Required=""{5}"" {6}/>";
        internal const string FIELD_XML_FORMAT_WITH_CHILD_NODES = @"<Field Type=""{0}"" Name=""{1}"" DisplayName=""{2}"" ID=""{3}"" Group=""{4}"" Required=""{5}"" {6}>{7}</Field>";
        internal const string FIELD_XML_PARAMETER_FORMAT = @"{0}=""{1}""";
        internal const string FIELD_XML_PARAMETER_WRAPPER_FORMAT = @"<Params {0}></Params>";
        internal const string FIELD_XML_USER_LISTIDENTIFIER = "UserInfo";
        internal const string FIELD_XML_USER_LISTRELATIVEURL = "_catalogs/users";
        internal const string FIELD_XML_CHILD_NODE = @"<{0}>{1}</{0}>";


        internal const string TAXONOMY_FIELD_XML_FORMAT = "<Field Type=\"{0}\" DisplayName=\"{1}\" ID=\"{8}\" ShowField=\"Term1033\" Required=\"{2}\" EnforceUniqueValues=\"FALSE\" {3} Sortable=\"FALSE\" Name=\"{4}\" Group=\"{9}\"><Default/><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value xmlns:q1=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q1:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">{5}</Value></Property><Property><Name>GroupId</Name></Property><Property><Name>TermSetId</Name><Value xmlns:q2=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q2:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">{6}</Value></Property><Property><Name>AnchorId</Name><Value xmlns:q3=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q3:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">00000000-0000-0000-0000-000000000000</Value></Property><Property><Name>UserCreated</Name><Value xmlns:q4=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q4:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>Open</Name><Value xmlns:q5=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q5:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>TextField</Name><Value xmlns:q6=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q6:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">{7}</Value></Property><Property><Name>IsPathRendered</Name><Value xmlns:q7=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q7:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">true</Value></Property><Property><Name>IsKeyword</Name><Value xmlns:q8=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q8:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>TargetTemplate</Name></Property><Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q9:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>FilterAssemblyStrongName</Name><Value xmlns:q10=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q10:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">Microsoft.SharePoint.Taxonomy, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Value></Property><Property><Name>FilterClassName</Name><Value xmlns:q11=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q11:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">Microsoft.SharePoint.Taxonomy.TaxonomyField</Value></Property><Property><Name>FilterMethodName</Name><Value xmlns:q12=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q12:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">GetFilteringHtml</Value></Property><Property><Name>FilterJavascriptProperty</Name><Value xmlns:q13=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q13:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">FilteringJavascript</Value></Property></ArrayOfProperty></Customization></Field>";
        internal const string THEMES_DIRECTORY = "/_catalogs/theme/15/{0}";
        internal const string MASTERPAGE_SEATTLE = "/_catalogs/masterpage/seattle.master";
        internal const string MASTERPAGE_DIRECTORY = "/_catalogs/masterpage/{0}";
        internal const string MASTERPAGE_CONTENT_TYPE = "0x01010500B45822D4B60B7B40A2BFCC0995839404";
        internal const string PAGE_LAYOUT_CONTENT_TYPE = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811";
        internal const string HTMLPAGE_LAYOUT_CONTENT_TYPE = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110003D357F861E29844953D5CAA1D4D8A3B001EC1BD45392B7A458874C52A24C9F70B";
        internal const string ListItemDirField = "FileDirRef";
        internal const string ListItemFileNameField = "FileLeafRef";
        internal static readonly string[] DefaultViewPages = new string[] { "/Forms/AllItems.aspx", "/Forms/Thumbnails.aspx" };
        internal static readonly string[] SkipPathes = new string[] { "/_catalogs/", "/Style Library", "/IWConvertedForms", "/Translation Packages" };
        internal const string AllItemCamlQuery = "<View Scope='RecursiveAll'><ViewFields><FieldRef Name='{0}'/><FieldRef Name='{1}'/></ViewFields></View>";

        public const string MINIMUMZONEIDREQUIREDSERVERVERSION = "16.0.4803.1200";

        internal static readonly uint[] SupportedLCIDs = { 1031, 1036, 2108, 1057, 1044, 1049, 2052, 1028, 1081, 1086, 1060, 1030, 1069, 1035, 1043, 1051, 1068, 1026, 1110, 1055, 1106, 1050, 1038, 1042, 1063, 1071, 1033, 1025, 1041, 1062, 1164, 1046, 2070, 9242, 1054, 5146, 1029, 3082, 1037, 1045, 10266, 2074, 1058, 1032, 1061, 1040, 1087, 1053, 1066, 1027, 1048 };

        internal const string ModernAudienceTargetingInternalName = "_ModernAudienceTargetUserField";
        internal const string ModernAudienceTargetingMultiLookupInternalName = "_ModernAudienceAadObjectIds";
        internal const string ClassicAudienceTargetingInternalName = "Target_x0020_Audiences";
    }
}
