﻿using System;
using System.Collections.Generic;

namespace PnP.Framework.Enums
{
    public static class BuiltInFieldId
    {
        /// <summary>
        /// Returns a GUID that represents the content type identifier of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid ContentTypeId = new Guid("{03e45e84-1992-4d42-9116-26f756012634}");
        /// <summary>
        /// Returns a GUID that represents the content type of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid ContentType = new Guid("{c042a256-787d-4a6f-8a8a-cf6ab767f12d}");
        public static readonly Guid ID = new Guid("{1d22ea11-1e32-424e-89ab-9fedbadb6ce1}");
        /// <summary>
        /// Returns a GUID that represents the last modified date and time information that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Modified = new Guid("{28cf69c5-fa48-462a-b5cd-27b6f9d2bd5f}");
        /// <summary>
        /// Returns a GUID that represents the date and time when the specified Windows SharePoint Services object was created.
        /// </summary>
        public static readonly Guid Created = new Guid("{8c06beca-0777-48f7-91c7-6da68bc07b69}");
        /// <summary>
        /// Returns a GUID that represents the specified author of the Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Author = new Guid("{1df5e554-ec7e-46a6-901d-d85a3881cb18}");
        /// <summary>
        /// Returns a GUID that is used to represent the editor name or information that is associated with a person who is referenced by a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Editor = new Guid("{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}");
        /// <summary>
        /// Returns a GUID that represents the internal version of a rowset, and is used when the rowset includes multiple versions of the same Windows SharePoint Services list item object.
        /// </summary>
        public static readonly Guid owshiddenversion = new Guid("{d4e44a66-ee3a-4d02-88c9-4ec5ff3f4cd5}");
        public static readonly Guid Subject = new Guid("{76a81629-44d4-4ce1-8d4d-6d7ebcd885fc}");
        /// <summary>
        /// Returns a GUID that represents the author of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _Author = new Guid("{246d0907-637c-46b7-9aa0-0bb914daa832}");
        /// <summary>
        /// Returns a string that represents the category of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _Category = new Guid("{0fc9cace-c5c2-465d-ae88-b67f2964ca93}");
        public static readonly Guid _Status = new Guid("{1dab9b48-2d1a-47b3-878c-8e84f0d211ba}");
        /// <summary>
        /// Returns a GUID that represents information about the server-relative URL for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid FileRef = new Guid("{94f89715-e097-4e8b-ba79-ea02aa8b7adb}");
        /// <summary>
        /// Returns a GUID that represents information about the file directory for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid FileDirRef = new Guid("{56605df6-8fa1-47e4-a04c-5b384d59609f}");
        /// <summary>
        /// Returns a GUID that represents version control information for the last modified version of the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid Last_x0020_Modified = new Guid("{173f76c8-aebd-446a-9bc9-769a2bd2c18f}");
        /// <summary>
        /// Returns a GUID that indicates the date that is associated with the creation of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Created_x0020_Date = new Guid("{998b5cff-4a35-47a7-92f3-3914aa6aa4a2}");
        /// <summary>
        /// Returns a GUID that represents information about the file size for the version history of the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid File_x0020_Size = new Guid("{8fca95c0-9b7d-456f-8dae-b41ee2728b85}");
        /// <summary>
        /// Returns a GUID that represents information about the file system type for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid FSObjType = new Guid("{30bb605f-5bae-48fe-b4e3-1f81d9772af9}");
        public static readonly Guid PermMask = new Guid("{ba3c27ee-4791-4867-8821-ff99000bac98}");
        /// <summary>
        /// Returns a GUID that represents the designated user who has checked out the Windows SharePoint Services object by using version control.
        /// </summary>
        public static readonly Guid CheckoutUser = new Guid("{3881510a-4e4a-4ee8-b102-8ee8e2d0dd4b}");
        /// <summary>
        /// Returns a GUID that represents the virus scan status of a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid VirusStatus = new Guid("{4a389cb9-54dd-4287-a71a-90ff362028bc}");
        /// <summary>
        /// Returns a GUID that represents the associated instance identifier for the history of the specified Windows SharePoint Services list object that was used under version control.
        /// </summary>
        public static readonly Guid InstanceID = new Guid("{50a54da4-1528-4e67-954a-e2d24f1e9efb}");
        /// <summary>
        /// Returns a GUID that represents the check-in comments of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _CheckinComment = new Guid("{58014f77-5463-437b-ab67-eec79532da67}");
        public static readonly Guid MetaInfo = new Guid("{687c7f94-686a-42d3-9b67-2782eac4b4f8}");
        /// <summary>
        /// Returns a GUID that represents the version control alert level information of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _Level = new Guid("{43bdd51b-3c5b-4e78-90a8-fb2087f71e70}");
        /// <summary>
        /// Returns a GUID that represents the latest version information of the check-in history of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _IsCurrentVersion = new Guid("{c101c3e7-122d-4d4d-bc34-58e94a38c816}");
        /// <summary>
        /// Returns a GUID that represents the destination information of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _HasCopyDestinations = new Guid("{26d0756c-986a-48a7-af35-bf18ab85ff4a}");
        /// <summary>
        /// Returns a GUID that represents the CopySource property of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _CopySource = new Guid("{6b4e226d-3d88-4a36-808d-a129bf52bccf}");
        /// <summary>
        /// Returns a GUID that represents the information about the moderation status of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _ModerationStatus = new Guid("{fdc3b2ed-5bf2-4835-a4bc-b885f3396a61}");
        /// <summary>
        /// Returns a GUID that represents information about the moderation comments  of the specified Windows SharePoint Services weblog object.
        /// </summary>
        public static readonly Guid _ModerationComments = new Guid("{34ad21eb-75bd-4544-8c73-0e08330291fe}");
        /// <summary>
        /// Returns a GUID that represents information about the occupational title of a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Title = new Guid("{fa564e0f-0c70-4ab9-b863-0177e6ddd247}");
        public static readonly Guid WorkflowVersion = new Guid("{f1e020bc-ba26-443f-bf2f-b68715017bbc}");
        /// <summary>
        /// Returns a GUID that represents the attachments that are associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Attachments = new Guid("{67df98f4-9dec-48ff-a553-29bece9c5bf4}");
        /// <summary>
        /// Returns a GUID that indicates the editing state icon that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Edit = new Guid("{503f1caa-358e-4918-9094-4a2cdc4bc034}");
        public static readonly Guid LinkTitleNoMenu = new Guid("{bc91a437-52e7-49e1-8c4e-4698904b2b6d}");
        public static readonly Guid LinkTitle = new Guid("{82642ec8-ef9b-478f-acf9-31f7d45fbc31}");
        public static readonly Guid SelectTitle = new Guid("{b1f7969b-ea65-42e1-8b54-b588292635f2}");
        public static readonly Guid Order = new Guid("{ca4addac-796f-4b23-b093-d2a3f65c0774}");
        /// <summary>
        /// Returns a GUID that is used to return the GUID of the specified Windows SharePoint Services object.
        /// </summary>
        #pragma warning disable CA1720
        public static readonly Guid GUID = new Guid("{ae069f25-3ac2-4256-b9c3-15dbc15da0e0}");
        #pragma warning restore CA1720
        /// <summary>
        /// Returns a GUID that represents the workflow instance identifier that is specified in a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowInstanceID = new Guid("{de8beacf-5505-47cd-80a6-aa44e7ffe2f4}");
        public static readonly Guid UniqueId = new Guid("{4b7403de-8d94-43e8-9f0f-137a3e298126}");
        public static readonly Guid ProgId = new Guid("{c5c4b81c-f1d9-4b43-a6a2-090df32ebb68}");
        /// <summary>
        /// Returns a GUID that represents information about the server-relative URL for the file node that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid FileLeafRef = new Guid("{8553196d-ec8d-4564-9861-3dbe931050c8}");
        public static readonly Guid ScopeId = new Guid("{dddd2420-b270-4735-93b5-92b713d0944d}");
        /// <summary>
        /// Returns a GUID that represents information about the fully qualified e-mail sender for the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid EmailSender = new Guid("{4ce600fb-a927-4911-bfc1-11076b76b522}");
        /// <summary>
        /// Returns a GUID that represents the "sent to" information for the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid EmailTo = new Guid("{caa2cb1e-a124-4068-9496-14feef1a901f}");
        /// <summary>
        /// Returns a GUID that is used to represent information about the e-mail carbon copy recipient for the specified SharePoint list object.
        /// </summary>
        public static readonly Guid EmailCc = new Guid("{a6af6df4-feb5-4dbf-bef6-d81230d4a071}");
        /// <summary>
        /// Returns a GUID that is used to represent information about the e-mail display name for the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid EmailFrom = new Guid("{e7cb6f60-f676-4b1d-89a3-975b6bc78cad}");
        /// <summary>
        /// Returns a GUID that represents the subject information for the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid EmailSubject = new Guid("{072e9bb6-a643-44ce-b6fb-8b574a792556}");
        /// <summary>
        /// Returns a GUID that is used to represent the update identifier that is associated with the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid EmailCalendarUid = new Guid("{f4e00567-8a9d-451b-82d4-a4447f9bd9a5}");
        /// <summary>
        /// Returns a GUID that is used to represent the sequence modification number that is associated with the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid EmailCalendarSequence = new Guid("{7a0cb12b-c70c-4f99-99f1-a232783a87d7}");
        /// <summary>
        /// Returns a GUID that is used to represent the date stamp information that is associated with the specified Windows SharePoint Services calendar object.
        /// </summary>
        public static readonly Guid EmailCalendarDateStamp = new Guid("{32f182ba-284e-4a87-93c3-936a6585af39}");
        /// <summary>
        /// Returns a GUID that represents the version number of the user interface of the specified Windows SharePoint Services weblog object.
        /// </summary>
        public static readonly Guid _UIVersion = new Guid("{7841bf41-43d0-4434-9f50-a673baef7631}");
        /// <summary>
        /// Returns a GUID that represents the version string that is associated with the user interface of the specified Windows SharePoint Services weblog object.
        /// </summary>
        public static readonly Guid _UIVersionString = new Guid("{dce8262a-3ae9-45aa-aab4-83bd75fb738a}");
        public static readonly Guid Modified_x0020_By = new Guid("{822c78e3-1ea9-4943-b449-57863ad33ca9}");
        /// <summary>
        /// Returns a GUID that indicates the user who is associated with the creation of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Created_x0020_By = new Guid("{4dd7e525-8d6b-4cb4-9d3e-44ee25f973eb}");
        /// <summary>
        /// Returns a GUID that represents file type information that is associated with the version history for the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid File_x0020_Type = new Guid("{39360f11-34cf-4356-9945-25c44e68dade}");
        public static readonly Guid HTML_x0020_File_x0020_Type = new Guid("{0c5e0085-eb30-494b-9cdd-ece1d3c649a2}");
        /// <summary>
        /// Returns a GUID that represents the source URL of the specified Windows SharePoint Services weblog object.
        /// </summary>
        public static readonly Guid _SourceUrl = new Guid("{c63a459d-54ba-4ab7-933a-dcf1c6fadec2}");
        /// <summary>
        /// Returns a GUID that represents information about the shared file index of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _SharedFileIndex = new Guid("{034998e9-bf1c-4288-bbbd-00eacfc64410}");
        /// <summary>
        /// Returns a GUID that represents the icon that is used to create a link to a file in a document library, where the file can be edited without using a menu.
        /// </summary>
        public static readonly Guid LinkFilenameNoMenu = new Guid("{9d30f126-ba48-446b-b8f9-83745f322ebe}");
        /// <summary>
        /// Returns a GUID that represents the EditMenuTableEnd property of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _EditMenuTableStart = new Guid("{3c6303be-e21f-4366-80d7-d6d0a3b22c7a}");
        /// <summary>
        /// Returns a GUID that represents the EditMenuTableEnd property of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _EditMenuTableEnd = new Guid("{2ea78cef-1bf9-4019-960a-02c41636cb47}");
        /// <summary>
        /// Returns a GUID that represents the icon that is used to create a link to a file in a document library, where the file can be edited by using a menu.
        /// </summary>
        public static readonly Guid LinkFilename = new Guid("{5cc6dc79-3710-4374-b433-61cb4a686c12}");
        public static readonly Guid SelectFilename = new Guid("{5f47e085-2150-41dc-b661-442f3027f552}");
        /// <summary>
        /// Returns a GUID that specifies the document icon that is associated with the creation of the specified Windows SharePoint Services template document object.
        /// </summary>
        public static readonly Guid DocIcon = new Guid("{081c6e4c-5c14-4f20-b23e-1a71ceb6a67c}");
        public static readonly Guid ServerUrl = new Guid("{105f76ce-724a-4bba-aece-f81f2fce58f5}");
        /// <summary>
        /// Returns a GUID that represents the encoded search URL for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid EncodedAbsUrl = new Guid("{7177cfc7-f399-4d4d-905d-37dd51bc90bf}");
        /// <summary>
        /// Returns a GUID that represents the base name of the specified Windows SharePoint Services object that does not include URL information.
        /// </summary>
        public static readonly Guid BaseName = new Guid("{7615464b-559e-4302-b8e2-8f440b913101}");
        /// <summary>
        /// Returns a GUID that represents information about the properly formatted file size for version history of the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid FileSizeDisplay = new Guid("{78a07ba4-bda8-4357-9e0f-580d64487583}");
        /// <summary>
        /// Returns a GUID that represents the body of the specified Windows SharePoint Services message object.
        /// </summary>
        public static readonly Guid Body = new Guid("{7662cd2c-f069-4dba-9e35-082cf976e170}");
        /// <summary>
        /// Returns a GUID that represents expiration date for the specified Windows SharePoint Services announcement object.
        /// </summary>
        public static readonly Guid Expires = new Guid("{6a09e75b-8d17-4698-94a8-371eda1af1ac}");
        /// <summary>
        /// Returns a GUID that represents the URL of a Windows SharePoint Services link object.
        /// </summary>
        public static readonly Guid URL = new Guid("{c29e077d-f466-4d8e-8bbe-72b66c5f205c}");
        /// <summary>
        /// Returns a GUID that represents comments that are associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid _Comments = new Guid("{52578fc3-1f01-4f4d-b016-94ccbcf428cf}");
        public static readonly Guid _EndDate = new Guid("{8a121252-85a9-443d-8217-a1b57020fadf}");
        /// <summary>
        /// Returns a GUID that represents the end date for the specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid EndDate = new Guid("{2684f9f2-54be-429f-ba06-76754fc056bf}");
        /// <summary>
        /// Returns a GUID that represents the associated URL for the specified Windows SharePoint Services link object.
        /// </summary>
        public static readonly Guid URLwMenu = new Guid("{2a9ab6d3-268a-4c1c-9897-e5f018f87e64}");
        public static readonly Guid URLNoMenu = new Guid("{aeaf07ee-d2fb-448b-a7a3-cf7e062d6c2a}");
        /// <summary>
        /// Returns a GUID that represents information about the phonetics (speech sounds) that are associated with the last name of the specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid LastNamePhonetic = new Guid("{fdc8216d-dabf-441d-8ac0-f6c626fbdc24}");
        /// <summary>
        /// Returns a GUID that represents the first name for the specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid FirstName = new Guid("{4a722dd4-d406-4356-93f9-2550b8f50dd0}");
        /// <summary>
        /// Returns a GUID that represents the phonetic information (speech sounds) that are associated with the specified first name for the Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid FirstNamePhonetic = new Guid("{ea8f7ca9-2a0e-4a89-b8bf-c51a6af62c73}");
        /// <summary>
        /// Returns a GUID that represents information about the full name of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid FullName = new Guid("{475c2610-c157-4b91-9e2d-6855031b3538}");
        /// <summary>
        /// Returns a GUID that represents pronunciation information for the person or company that is specified in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid CompanyPhonetic = new Guid("{034aae88-6e9a-4e41-bc8a-09b6c15fcdf4}");
        /// <summary>
        /// Returns a GUID that represents the company information for the person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Company = new Guid("{038d1503-4629-40f6-adaf-b47d1ab2d4fe}");
        /// <summary>
        /// Returns a GUID that represents the job title of the person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid JobTitle = new Guid("{c4e0f350-52cc-4ede-904c-dd71a3d11f7d}");
        /// <summary>
        /// Returns a GUID that represents the corporate telephone number for the person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkPhone = new Guid("{fd630629-c165-4513-b43c-fdb16b86a14d}");
        /// <summary>
        /// Returns a GUID that represents the home telephone number for the person specified in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomePhone = new Guid("{2ab923eb-9880-4b47-9965-ebf93ae15487}");
        /// <summary>
        /// Returns a GUID that represents the cell phone number of the person who is specified in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid CellPhone = new Guid("{2a464df1-44c1-4851-949d-fcd270f0ccf2}");
        /// <summary>
        /// Returns a GUID that represents the corporate fax information for the person specified in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkFax = new Guid("{9d1cacc8-f452-4bc1-a751-050595ad96e1}");
        /// <summary>
        /// Returns a GUID that represents the work address of the person who is referenced in the Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkAddress = new Guid("{fc2e188e-ba91-48c9-9dd3-16431afddd50}");
        public static readonly Guid _Photo = new Guid("{1020c8a0-837a-4f1b-baa1-e35aff6da169}");
        /// <summary>
        /// Returns a GUID that represents the work city for the person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkCity = new Guid("{6ca7bd7f-b490-402e-af1b-2813cf087b1e}");
        /// <summary>
        /// Returns a GUID that represents the regional corporate information for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkState = new Guid("{ceac61d3-dda9-468b-b276-f4a6bb93f14f}");
        /// <summary>
        /// Returns a GUID that represents the work ZIP code for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkZip = new Guid("{9a631556-3dac-49db-8d2f-fb033b0fdc24}");
        /// <summary>
        /// Returns a GUID that represents the corporate country information for a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WorkCountry = new Guid("{3f3a5c85-9d5a-4663-b925-8b68a678ea3a}");
        /// <summary>
        /// Returns a GUID that represents the Web page that is associated with a person who is referenced in a specified  Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid WebPage = new Guid("{a71affd2-dcc7-4529-81bc-2fe593154a5f}");
        /// <summary>
        /// Returns a GUID that represents the priority information that is associated with a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid Priority = new Guid("{a8eb573e-9e11-481a-a8c9-1104a54b2fbd}");
        /// <summary>
        /// Returns a GUID that represents information about the enumerated completion status for a specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid TaskStatus = new Guid("{c15b34c3-ce7d-490a-b133-3f4de8801b76}");
        /// <summary>
        /// Returns a GUID that represents information about what percent of a specified Windows SharePoint Services workflow object is completed.
        /// </summary>
        public static readonly Guid PercentComplete = new Guid("{d2311440-1ed6-46ea-b46d-daa643dc3886}");
        /// <summary>
        /// Returns a GUID that indicates the user to whom the specified Windows SharePoint Services workflow task object is assigned.
        /// </summary>
        public static readonly Guid AssignedTo = new Guid("{53101f38-dd2e-458c-b245-0c236cc13d1a}");
        /// <summary>
        /// Returns a GUID that represents information about the task group for a specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid TaskGroup = new Guid("{50d8f08c-8e99-4948-97bf-2be41fa34a0d}");
        /// <summary>
        /// Returns a GUID that represents information about the start date of a task that is associated with the specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid StartDate = new Guid("{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}");
        /// <summary>
        /// Returns a GUID that represents information about the due date for a specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid TaskDueDate = new Guid("{cd21b4c2-6841-4f9e-a23a-738a65f99889}");
        /// <summary>
        /// Returns a GUID that represents the workflow URL that is specified in a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowLink = new Guid("{58ddda52-c2a3-4650-9178-3bbc1f6e36da}");
        /// <summary>
        /// Returns a GUID that represents information about an off-site participant user object that is associated with the specified Windows SharePoint Services workflow object.
        /// </summary>
        public static readonly Guid OffsiteParticipant = new Guid("{16b6952f-3ce6-45e0-8f4e-42dac6e12441}");
        /// <summary>
        /// Returns a GUID that represents information about why an offsite participant is offsite. This information, in turn, is associated with the specified Windows SharePoint Services workflow object.
        /// </summary>
        public static readonly Guid OffsiteParticipantReason = new Guid("{4a799ba5-f449-4796-b43e-aa5186c3c414}");
        /// <summary>
        /// Returns a GUID that represents the type of outcome (for example, "Approved" or "Rejected") that is associated with a specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowOutcome = new Guid("{18e1c6fa-ae37-4102-890a-cfb0974ef494}");
        /// <summary>
        /// Returns a GUID that represents the workflow name that is specified in a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowName = new Guid("{e506d6ca-c2da-4164-b858-306f1c41c9ec}");
        /// <summary>
        /// Returns a GUID that represents information about the task type of a specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid TaskType = new Guid("{8d96aa48-9dff-46cf-8538-84c747ffa877}");
        public static readonly Guid FormURN = new Guid("{17ca3a22-fdfe-46eb-99b5-9646baed3f16}");
        public static readonly Guid FormData = new Guid("{78eae64a-f5f2-49af-b416-3247b76f46a1}");
        /// <summary>
        /// Returns a GUID that is used to represent information about the custom e-mail body of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid EmailBody = new Guid("{8cbb9252-1035-4156-9c35-f54e9056c65a}");
        /// <summary>
        /// Returns a GUID that references the custom e-mail body of the specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid HasCustomEmailBody = new Guid("{47f68c3b-8930-406f-bde2-4a8c669ee87c}");
        public static readonly Guid SendEmailNotification = new Guid("{cb2413f2-7de9-4afc-8587-1ca3f563f624}");
        /// <summary>
        /// Returns a GUID that represents the time when the specified Windows SharePoint Services workflow object modifications were accepted (and not rolled back to an earlier version).
        /// </summary>
        public static readonly Guid PendingModTime = new Guid("{4d2444c2-0e97-476c-a2a3-e9e4a9c73009}");
        /// <summary>
        /// Returns a GUID that represents information about the completed status of the specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid Completed = new Guid("{35363960-d998-4aad-b7e8-058dfe2c669e}");
        public static readonly Guid WorkflowListId = new Guid("{1bfee788-69b7-4765-b109-d4d9c31d1ac1}");
        public static readonly Guid WorkflowItemId = new Guid("{8e234c69-02b0-42d9-8046-d5f49bf0174f}");
        public static readonly Guid ExtendedProperties = new Guid("{1c5518e2-1e99-49fe-bfc6-1a8de3ba16e2}");
        public static readonly Guid AdminTaskAction = new Guid("{7b016ee5-70aa-4abb-8aa3-01795b4efe6f}");
        public static readonly Guid AdminTaskDescription = new Guid("{93490584-b6a8-4996-aa00-ead5f59aae0d}");
        public static readonly Guid AdminTaskOrder = new Guid("{cf935cc2-a00c-4ad3-bca1-0865ab15afc1}");
        public static readonly Guid Service = new Guid("{48b4a73e-8853-44ac-83a8-3a4bd59ce9ec}");
        public static readonly Guid SystemTask = new Guid("{af0a3d4b-3ceb-449e-9bf4-51103f2032e3}");
        /// <summary>
        /// Returns a GUID that represents information about the physical location that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Location = new Guid("{288f5f32-8462-4175-8f09-dd7ba29359a9}");
        /// <summary>
        /// Returns a GUID that represents information about the recurrence field of the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid fRecurrence = new Guid("{f2e63656-135e-4f1c-8fc2-ccbe74071901}");
        /// <summary>
        /// Returns a GUID that represents the URL for the meeting workspace for a specified Windows SharePoint Services event object.
        /// </summary>
        public static readonly Guid WorkspaceLink = new Guid("{08fc65f9-48eb-4e99-bd61-5946c439e691}");
        /// <summary>
        /// Returns a GUID that represents information about the event type for the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid EventType = new Guid("{5d1d4e76-091a-4e03-ae83-6a59847731c0}");
        /// <summary>
        /// Returns a GUID that represents an update identifier for a Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid UID = new Guid("{63055d04-01b5-48f3-9e1e-e564e7c6b23b}");
        /// <summary>
        /// Returns a GUID that represents the recurrence identifier for a specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid RecurrenceID = new Guid("{dfcc8fff-7c4c-45d6-94ed-14ce0719efef}");
        /// <summary>
        /// Returns a GUID that represents cancellation information for the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid EventCanceled = new Guid("{b8bbe503-bb22-4237-8d9e-0587756a2176}");
        /// <summary>
        /// Returns a GUID that is associated with the duration of an event, as represented in a Windows SharePoint Services workflow event object.
        /// </summary>
        public static readonly Guid Duration = new Guid("{4d54445d-1c84-4a6d-b8db-a51ded4e1acc}");
        /// <summary>
        /// Returns a GUID that represents information about the recurrence data that is associated with a Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid RecurrenceData = new Guid("{d12572d0-0a1e-4438-89b5-4d0430be7603}");
        /// <summary>
        /// Returns a GUID that represents the time zone that is associated with a specified Windows SharePoint Services site or user object.
        /// </summary>
        public static readonly Guid TimeZone = new Guid("{6cc1c612-748a-48d8-88f2-944f477f301b}");
        /// <summary>
        /// Returns a GUID that indicates a time zone that is expressed in XML format and is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid XMLTZone = new Guid("{c4b72ed6-45aa-4422-bff1-2b6750d30819}");
        public static readonly Guid MasterSeriesItemID = new Guid("{9b2bed84-7769-40e3-9b1d-7954a4053834}");
        /// <summary>
        /// Returns a GUID that represents the meeting workspace information for a specified Windows SharePoint Services event object.
        /// </summary>
        public static readonly Guid Workspace = new Guid("{881eac4a-55a5-48b6-a28e-8329d7486120}");
        /// <summary>
        /// Returns a GUID that is used to indicate the issue status (for example, "Active", "Resolved", "Closed") that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid IssueStatus = new Guid("{3f277a5c-c7ae-4bbe-9d44-0456fb548f94}");
        /// <summary>
        /// Returns a GUID that represents the comments that are associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Comment = new Guid("{6df9bd52-550e-4a30-bc31-a4366832a87f}");
        /// <summary>
        /// Returns a GUID that represents comments that are associated with a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Comments = new Guid("{9da97a8a-1da5-4a77-98d3-4bc10456e700}");
        /// <summary>
        /// Returns a GUID that represents the data category of the specified Windows SharePoint Services message object.
        /// </summary>
        public static readonly Guid Category = new Guid("{6df9bd52-550e-4a30-bc31-a4366832a87d}");
        /// <summary>
        /// Returns a GUID that represents related issues for a specified Windows SharePoint Services issue object.
        /// </summary>
        public static readonly Guid RelatedIssues = new Guid("{875fab27-6e95-463b-a4a6-82544f1027fb}");
        public static readonly Guid LinkIssueIDNoMenu = new Guid("{03f89857-27c9-4b58-aaab-620647deda9b}");
        /// <summary>
        /// Returns a GUID that represents the comments that are associated with a change in issue status (for example, changing an item to "Active", "Resolved", or "Closed" status) of the specified Windows SharePoint Services issue object.
        /// </summary>
        public static readonly Guid V3Comments = new Guid("{6df9bd52-550e-4a30-bc31-a4366832a87e}");
        /// <summary>
        /// Returns a GUID that represents the name of the person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Name = new Guid("{bfc6f32c-668c-43c4-a903-847cca2f9b3c}");
        /// <summary>
        /// Returns a GUID that is used to represent the e-mail address of a person who is represented by a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid EMail = new Guid("{fce16b4c-fe53-4793-aaab-b4892e736d15}");
        /// <summary>
        /// Returns a GUID that represents the notes that are associated with the person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Notes = new Guid("{e241f186-9b94-415c-9f66-255ce7f86235}");
        /// <summary>
        /// Returns a GUID that indicates whether the person who is associated with the specified Windows SharePoint Services user object is a site administrator.
        /// </summary>
        public static readonly Guid IsSiteAdmin = new Guid("{9ba260b2-85a1-4a32-ad7a-63eaceffe6b4}");
        /// <summary>
        /// Returns a GUID that is used to represent deletion information that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Deleted = new Guid("{4ed6dfdf-86a8-4894-bd1b-4fa28042be53}");
        /// <summary>
        /// Returns a GUID that represents the graphic image that is associated with a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Picture = new Guid("{d9339777-b964-489a-bf09-2ac3c3fe5f0d}");
        /// <summary>
        /// Returns a GUID that is used to represent information about the department information for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Department = new Guid("{05fdf852-4b64-4096-9b2b-d2a62a86bc59}");
        /// <summary>
        /// Returns a GUID that represents the Session Initiation Protocol (SIP) information for a Windows SharePoint Services user object.
        /// </summary>
        public static readonly Guid SipAddress = new Guid("{829c275d-8744-4d9b-a42f-53f53eb60559}");
        /// <summary>
        /// Returns a GUID that is used to indicate whether a person who is associated with the specified Windows SharePoint Services user profile object is marked as active or inactive by the site administrator.
        /// </summary>
        public static readonly Guid IsActive = new Guid("{af5036db-36f4-46c8-bde7-a677bd0ef280}");
        /// <summary>
        /// Returns a GUID that represents whether or not the text that follows the first HTML division has been trimmed for the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid TrimmedBody = new Guid("{6d0f8993-5050-41f3-be6c-18902d282357}");
        public static readonly Guid DiscussionLastUpdated = new Guid("{59956c56-30dd-4cb1-bf12-ef693b42679c}");
        public static readonly Guid MessageId = new Guid("{2ef29342-2f5f-4052-90d3-8192e0705e51}");
        public static readonly Guid ThreadTopic = new Guid("{769b99d9-d361-4948-b687-f01332391629}");
        public static readonly Guid ThreadIndex = new Guid("{cef73bf1-edf6-4dd9-9098-a07d83984700}");
        /// <summary>
        /// Returns a GUID that represents information about the header of the specified Windows SharePoint Services list object.
        /// </summary>
        public static readonly Guid EmailReferences = new Guid("{124527a9-fc10-48ff-8d44-960a7db405f8}");
        public static readonly Guid RelevantMessages = new Guid("{9161f6cb-a8e6-47b8-9d24-89415de691f7}");
        /// <summary>
        /// Returns a GUID that represents the identifier of the parent folder of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid ParentFolderId = new Guid("{a9ec25bf-5a22-4658-bd19-484e52efbe1a}");
        public static readonly Guid ShortestThreadIndex = new Guid("{4753e73b-5b5d-4bbc-8e09-c9683b0d40a7}");
        public static readonly Guid ShortestThreadIndexId = new Guid("{2bec4782-695f-406d-9e50-f1d39a2b8eb6}");
        public static readonly Guid ShortestThreadIndexIdLookup = new Guid("{8ffccefe-998b-4896-a6df-32d566f69141}");
        /// <summary>
        /// Returns a GUID that is used to represent lookup information for the associated title of the specified Windows SharePoint Services discussion board message object.
        /// </summary>
        public static readonly Guid DiscussionTitleLookup = new Guid("{f0218b98-d0d6-4fc1-b15b-aabeb89f32a9}");
        /// <summary>
        /// Returns a GUID that is used to represent information about the discussion title for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid DiscussionTitle = new Guid("{c5abfdc7-3435-4183-9207-3d1146895cf8}");
        public static readonly Guid ParentItemEditor = new Guid("{ff90fecb-7f46-44f5-9698-db44a81b2a8b}");
        public static readonly Guid ParentItemID = new Guid("{7d014138-1886-41f0-834f-ba9f4e72f33b}");
        public static readonly Guid LastReplyBy = new Guid("{7f15088c-1448-41c7-a125-18a3a90ce543}");
        public static readonly Guid IsQuestion = new Guid("{7aead996-f9f9-4682-9e0e-f5634ab352c8}");
        public static readonly Guid BestAnswerId = new Guid("{a8b93fba-7396-443d-9884-ee332caa4560}");
        public static readonly Guid IsAnswered = new Guid("{32b1ca82-a25b-48d1-b78d-3a956ba07c41}");
        /// <summary>
        /// Returns a GUID that represents the linked discussion thread title, in which the discussion items in a discussion board can be edited without using  a menu.
        /// </summary>
        public static readonly Guid LinkDiscussionTitleNoMenu = new Guid("{3ac9353f-613f-42bd-98e1-530e9fd1cbf6}");
        /// <summary>
        /// Returns a GUID that represents the linked discussion thread title, in which discussion items in a discussion board can be edited by using a menu.
        /// </summary>
        public static readonly Guid LinkDiscussionTitle = new Guid("{46045bc4-283a-4826-b3dd-7a78d790b266}");
        public static readonly Guid LinkDiscussionTitle2 = new Guid("{b4e31c47-f962-4f9f-9132-eb555a1a026c}");
        public static readonly Guid ReplyNoGif = new Guid("{87cda0e2-fc57-4eec-a696-b0de2f61f361}");
        public static readonly Guid ThreadingControls = new Guid("{c55a4674-640b-4bae-8738-ce0439e6f6d4}");
        /// <summary>
        /// Returns a GUID that represents information about the associated reply indentation level of the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid IndentLevel = new Guid("{68227570-72dd-4816-b6b6-4b81ff99a393}");
        /// <summary>
        /// Returns a GUID that represents information about the associated reply indentation level of the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid Indentation = new Guid("{26c4f53e-733a-4202-814b-377492b6c841}");
        public static readonly Guid StatusBar = new Guid("{f90bce56-87dc-4d73-bfcb-03fcaf670500}");
        /// <summary>
        /// Returns a GUID that represents the body and associated indexing information (for example, the subject or discussion thread title) of the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid BodyAndMore = new Guid("{c7e9537e-bde4-4923-a100-adbd9e0a0a0d}");
        public static readonly Guid MessageBody = new Guid("{fbba993f-afee-4e00-b9be-36bc660dcdd1}");
        /// <summary>
        /// Returns a GUID that represents the expansion of the body and associated indexing information (for example, the subject or discussion thread title) of the specified Windows SharePoint Services message object.
        /// </summary>
        public static readonly Guid BodyWasExpanded = new Guid("{af82aa75-3039-4573-84a8-73ffdfd22733}");
        public static readonly Guid QuotedTextWasExpanded = new Guid("{e393d344-2e8c-425b-a8c3-89ac3144c9a2}");
        /// <summary>
        /// Returns a GUID that represents the appropriate message body to display.
        /// </summary>
        public static readonly Guid CorrectBodyToShow = new Guid("{b0204f69-2253-43d2-99ad-c0df00031b66}");
        /// <summary>
        /// Returns a GUID that represents a link that provides a large amount of disclosed information from the message body for the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid FullBody = new Guid("{9c4be348-663a-4172-a38a-9714b2634c17}");
        /// <summary>
        /// Returns a GUID that represents a link that provides a small amount of disclosed information from the message body for the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid LimitedBody = new Guid("{61b97279-cbc0-4aa9-a362-f1ff249c1706}");
        /// <summary>
        /// Returns a GUID that represents a link that provides a large amount of disclosed information (in snippet form) for the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid MoreLink = new Guid("{fb6c2494-1b14-49b0-a7ca-0506d6e85a62}");
        /// <summary>
        /// Returns a GUID that represents a link that provides a small amount of disclosed information (in snippet form) for the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid LessLink = new Guid("{076193bd-865b-4de7-9633-1f12069a6fff}");
        /// <summary>
        /// Returns a GUID that represents whether or not quoted text can be toggled in the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid ToggleQuotedText = new Guid("{e451420d-4e62-43e3-af83-010d36e353a2}");
        public static readonly Guid Threading = new Guid("{58ca6516-51cd-41fb-a908-dd2a4aeea8bc}");
        /// <summary>
        /// Returns a GUID that represents the image of the person who is referenced by a specified Windows SharePoint Services user object.
        /// </summary>
        public static readonly Guid PersonImage = new Guid("{adfe65ee-74bb-4771-bec5-d691d9a6a14e}");
        /// <summary>
        /// Returns a GUID that represents the minimal personal view of information that is associated with a user (for example, the header information that is associated with a discussion thread title) of the specified Windows SharePoint Services discussion board object.
        /// </summary>
        public static readonly Guid PersonViewMinimal = new Guid("{b4ab471e-0262-462a-8b3f-c1dfc9e2d5fd}");
        public static readonly Guid IsRootPost = new Guid("{bd2216c1-a2f3-48c0-b21c-dc297d0cc658}");
        /// <summary>
        /// Returns a GUID that represents combined file information as well as a specific base name of a Windows SharePoint Services object. This file information can be specific to a UNC path, a URL, local directories, or local files.
        /// </summary>
        public static readonly Guid Combine = new Guid("{e52012a0-51eb-4c0c-8dfb-9b8a0ebedcb6}");
        public static readonly Guid RepairDocument = new Guid("{5d36727b-bcb2-47d2-a231-1f0bc63b7439}");
        public static readonly Guid ShowRepairView = new Guid("{11851948-b05e-41be-9d9f-bc3bf55d1de3}");
        public static readonly Guid ShowCombineView = new Guid("{086f2b30-460c-4251-b75a-da88a5b205c1}");
        public static readonly Guid TemplateUrl = new Guid("{4b1bf6c6-4f39-45ac-acd5-16fe7a214e5e}");
        public static readonly Guid xd_ProgID = new Guid("{cd1ecb9f-dd4e-4f29-ab9e-e9ff40048d64}");
        public static readonly Guid xd_Signature = new Guid("{fbf29b2d-cae5-49aa-8e0a-29955b540122}");
        /// <summary>
        /// Returns a GUID that represents the workflow instance that is specified in a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowInstance = new Guid("{de21c770-a12b-4f88-af4b-aeebd897c8c2}");
        /// <summary>
        /// Returns a GUID that represents an identifier that is associated with  another workflow object, as specified with a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowAssociation = new Guid("{8d426880-8d96-459b-ae48-e8b3836d8b9d}");
        /// <summary>
        /// Returns a GUID that represents the template that is associated with a Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid WorkflowTemplate = new Guid("{bfb1589e-2016-4b98-ae62-e91979c3224f}");
        public static readonly Guid List = new Guid("{f44e428b-61c8-4100-a911-a3a635f43bb5}");
        /// <summary>
        /// Returns a GUID that represents the item identifier for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Item = new Guid("{92b8e9d0-a11b-418f-bf1c-c44aaa73075d}");
        /// <summary>
        /// Returns a GUID that represents the user information that is associated with the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid User = new Guid("{5928ff1f-daa1-406c-b4a9-190485a448cb}");
        /// <summary>
        /// Returns a GUID that represents whether or not a workflow event has occurred for the specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid Occurred = new Guid("{5602dc33-a60a-4dec-bd23-d18dfcef861d}");
        /// <summary>
        /// Returns a GUID that represents the name of the specified Windows SharePoint Services event object.
        /// </summary>
        public static readonly Guid Event = new Guid("{20a1a5b1-fddf-4420-ac68-9701490e09af}");
        /// <summary>
        /// Returns a GUID that represents information about the permissions group of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Group = new Guid("{c86a2f7f-7680-4a0b-8907-39c4f4855a35}");
        /// <summary>
        /// Returns a GUID that represents the outcome that is associated with the specified Windows SharePoint Services workflow object.
        /// </summary>
        public static readonly Guid Outcome = new Guid("{dcde7b1f-918b-4ed5-819f-9798f8abac37}");
        /// <summary>
        /// Returns a GUID that is associated with an event duration, as represented in a Windows SharePoint Services workflow event object This GUID is limited to a maximum of 255 characters.
        /// </summary>
        public static readonly Guid DLC_Duration = new Guid("{80289bac-fd36-4848-b67a-bc8b5b621ec2}");
        /// <summary>
        /// Returns a GUID that is associated with an event description, as represented in a Windows SharePoint Services workflow event object. This GUID is limited to a maximum of 255 characters.
        /// </summary>
        public static readonly Guid DLC_Description = new Guid("{2fd53156-ff9d-4cc3-b0ac-fe8a7bc82283}");
        /// <summary>
        /// Returns a GUID that represents data that is associated with a Windows SharePoint Services workflow event object.
        /// </summary>
        public static readonly Guid Data = new Guid("{38269294-165e-448a-a6b9-f0e09688f3f9}");
        public static readonly Guid Purpose = new Guid("{8ee23f39-e2d1-4b46-8945-42386b24829d}");
        /// <summary>
        /// Returns a GUID that represents the type of interface connection that is used with the associated Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid ConnectionType = new Guid("{939dfb93-3107-44c6-a98f-dd88dca3f8cf}");
        /// <summary>
        /// Returns a GUID that represents information about the file type for version history of the specified Windows SharePoint Services library picture object.
        /// </summary>
        public static readonly Guid FileType = new Guid("{c53a03f3-f930-4ef2-b166-e0f2210c13c0}");
        /// <summary>
        /// Returns a GUID that represents the image size of the specified Windows SharePoint Services image object.
        /// </summary>
        public static readonly Guid ImageSize = new Guid("{922551b8-c7e0-46a6-b7e3-3cf02917f68a}");
        /// <summary>
        /// Returns a GUID that represents the width of the specified Windows SharePoint Services image object.
        /// </summary>
        public static readonly Guid ImageWidth = new Guid("{7e68a0f9-af76-404c-9613-6f82bc6dc28c}");
        /// <summary>
        /// Returns a GUID that represents the height of the specified Windows SharePoint Services image object.
        /// </summary>
        public static readonly Guid ImageHeight = new Guid("{1944c034-d61b-42af-aa84-647f2e74ca70}");
        /// <summary>
        /// Returns a GUID that represents information about the creation date of the specified Windows SharePoint Services image object.
        /// </summary>
        public static readonly Guid ImageCreateDate = new Guid("{a5d2f824-bc53-422e-87fd-765939d863a5}");
        /// <summary>
        /// Returns a GUID that represents the URL of the encoded thumbnail search image for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid EncodedAbsThumbnailUrl = new Guid("{b9e6f3ae-5632-4b13-b636-9d1a2bd67120}");
        /// <summary>
        /// Returns a GUID that represents the encoded Web image of the search URL for the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid EncodedAbsWebImgUrl = new Guid("{a1ca0063-779f-49f9-999c-a4a2e3645b07}");
        public static readonly Guid SelectedFlag = new Guid("{7ebf72ca-a307-4c18-9e5b-9d89e1dae74f}");
        /// <summary>
        /// Returns a GUID that represents information about the image of the specified Windows SharePoint Services picture library object.
        /// </summary>
        public static readonly Guid NameOrTitle = new Guid("{76d1cc87-56de-432c-8a2a-16e5ba5331b3}");
        public static readonly Guid RequiredField = new Guid("{de1baa4b-2117-473b-aa0c-4d824034142d}");
        /// <summary>
        /// Returns a GUID that represents information about the keyword summary of the specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Keywords = new Guid("{b66e9b50-a28e-469b-b1a0-af0e45486874}");
        /// <summary>
        /// Returns a GUID that represents the thumbnail image for a Windows SharePoint Services image object.
        /// </summary>
        public static readonly Guid Thumbnail = new Guid("{ac7bb138-02dc-40eb-b07a-84c15575b6e9}");
        public static readonly Guid Preview = new Guid("{bd716b26-546d-43f2-b229-62699581fa9f}");
        /// <summary>
        /// Returns a GUID that represents the selected decision status that is associated with a Windows SharePoint Services workflow event object.
        /// </summary>
        public static readonly Guid DecisionStatus = new Guid("{ac3a1092-34ad-42b2-8d47-a79d01d9f516}");
        /// <summary>
        /// Returns a GUID that represents the availability status of the specified Windows SharePoint Services object designated as an attendee.
        /// </summary>
        public static readonly Guid AttendeeStatus = new Guid("{3329f39d-70ed-4858-b8c8-c5237634bf08}");
        /// <summary>
        /// Returns a GUID that represents the field that indicates an all-day event for the specified Windows SharePoint Services calendar event object.
        /// </summary>
        public static readonly Guid fAllDayEvent = new Guid("{7d95d1f4-f5fd-4a70-90cd-b35abc9b5bc8}");
        /// <summary>
        /// Returns a GUID that represents information about the primary spoken and written language of a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Language = new Guid("{d81529e8-384c-4ca6-9c43-c86a256e6a44}");
        public static readonly Guid SurveyTitle = new Guid("{e6f528fb-2e22-483d-9c80-f2536acdc6de}");
        /// <summary>
        /// Returns a GUID that is associated with the description of the content type of the wiki (for example, a "How To" wiki content type description) of the specified Windows SharePoint Services wiki document object.
        /// </summary>
        public static readonly Guid WikiField = new Guid("{c33527b4-d920-4587-b791-45024d00068a}");
        public static readonly Guid PublishedDate = new Guid("{b1b53d80-23d6-e31b-b235-3a286b9f10ea}");
        public static readonly Guid PostCategory = new Guid("{38bea83b-350a-1a6e-f34a-93a6af31338b}");
        public static readonly Guid BaseAssociationGuid = new Guid("{e9359d15-261b-48f6-a302-01419a68d4de}");
        public static readonly Guid XomlUrl = new Guid("{566da236-762b-4a76-ad1f-b08b3c703fce}");
        public static readonly Guid RulesUrl = new Guid("{ad97fbac-70af-4860-a078-5ee704946f93}");
        /// <summary>
        /// Returns a GUID that represents the categories that are associated with a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Categories = new Guid("{9ebcd900-9d05-46c8-8f4d-e46e87328844}");
        /// <summary>
        /// Returns a GUID that represents the address of an event that is represented by a specified Windows SharePoint Services event object.
        /// </summary>
        public static readonly Guid ol_EventAddress = new Guid("{493896da-0a4f-46ec-a68e-9cfd1a5fc19b}");
        /// <summary>
        /// Returns a GUID that represents the completion date that is associated with a specified Windows SharePoint Services task object.
        /// </summary>
        public static readonly Guid DateCompleted = new Guid("{24bfa3c2-e6a0-4651-80e9-3db44bf52147}");
        /// <summary>
        /// Returns a GUID that represents the total hours of work performed by a person or resource that is referenced by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid TotalWork = new Guid("{f3c4a259-19a2-44b8-ab3d-e9145d07d538}");
        /// <summary>
        /// Returns a GUID that represents the actual work value that is associated with a specified Windows SharePoint Services workflow task object.
        /// </summary>
        public static readonly Guid ActualWork = new Guid("{b0b3407e-1c33-40ed-a37c-2430b7a5d081}");
        public static readonly Guid TaskCompanies = new Guid("{3914f98e-6d99-4218-9ba3-af7370b9e7bc}");
        /// <summary>
        /// Returns a GUID that represents mileage information that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Mileage = new Guid("{3126c2f1-063e-4892-828f-0696ec6e105f}");
        /// <summary>
        /// Returns a GUID that represents the billing information that is associated with a person who is referenced by a specified  Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid BillingInformation = new Guid("{4f03f66b-fb1e-4ed2-ab8e-f6ed3fe14844}");
        /// <summary>
        /// Returns a GUID that represents an organizational role description for a person who is referenced by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Role = new Guid("{eeaeaaf1-4110-465b-905e-df1073a7e0e6}");
        /// <summary>
        /// Returns a GUID that represents the middle name of a person who is represented by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid MiddleName = new Guid("{418c8d29-6f2e-44c3-8955-2cd7ec3e2151}");
        /// <summary>
        /// Returns a GUID that represents the suffix for a person who is referenced by a specified Windows SharePoint Services contact object (such as M.D., Jr., Sr., or III).
        /// </summary>
        public static readonly Guid Suffix = new Guid("{d886eba3-d018-4103-a322-d5780127ef8a}");
        /// <summary>
        /// Returns a GUID that represents the telephone number of the assistant for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid AssistantNumber = new Guid("{f55de332-074e-4e71-a71a-b90abfad51ae}");
        /// <summary>
        /// Returns a GUID that represents a second corporate telephone number that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Business2Number = new Guid("{6547d03a-76d3-4d74-9d34-f51b837c0879}");
        /// <summary>
        /// Returns a GUID that represents a callback telephone number that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid CallbackNumber = new Guid("{344e9657-b17f-4344-a834-ff7c056bcc5e}");
        /// <summary>
        /// Returns a GUID that represents the car identification number that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid CarNumber = new Guid("{92a011a9-fd1b-42e0-b6fa-afcfee1928fa}");
        /// <summary>
        /// Returns a GUID that represents the main telephone number of a corporation that is associated with a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid CompanyNumber = new Guid("{27cb1283-bda2-4ae8-bcff-71725b674dbb}");
        /// <summary>
        /// Returns a GUID that represents the second home telephone number of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Home2Number = new Guid("{8c5a385d-2fff-42da-a4c5-f6a904f2e491}");
        /// <summary>
        /// Returns a GUID that represents the home facsimile telephone number of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeFaxNumber = new Guid("{c189a857-e6b0-488f-83a0-f4ee0a3ad01e}");
        /// <summary>
        /// Returns a GUID that represents the Integrated Services Digital Network (ISDN) number of a person who is represented by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid ISDNNumber = new Guid("{a579062a-6c1d-4ad3-9d5e-035f9f2c1882}");
        /// <summary>
        /// Returns a GUID that represents an alternative telephone number for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherNumber = new Guid("{96e02495-f428-48bc-9f13-06d98ba58c34}");
        /// <summary>
        /// Returns a GUID that represents an alternative facsimile telephone number for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherFaxNumber = new Guid("{aad15eb6-d7fd-47b8-abd4-adc0fe33a6ba}");
        /// <summary>
        /// Returns a GUID that represents the number of a pager device for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid PagerNumber = new Guid("{f79bf074-daf7-4c06-a314-15b287fdf4c9}");
        /// <summary>
        /// Returns a GUID that represents the primary telephone number of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid PrimaryNumber = new Guid("{d69bcc0e-57c3-4f3b-bbc5-b090edf21f0f}");
        /// <summary>
        /// Returns a GUID that represents the number of a portable radio unit of a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid RadioNumber = new Guid("{d1aede4f-1352-48d9-81e2-b10097c359c1}");
        /// <summary>
        /// Returns a GUID that represents the Telex number of a person who is represented by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid TelexNumber = new Guid("{e7be7f3c-c436-481d-8865-669e5146f53c}");
        /// <summary>
        /// Returns a GUID that represents the number of a TeleType (TTY) or Telephone Device for the Deaf (TDD) of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid TTYTDDNumber = new Guid("{f54697f1-0357-4c5a-a711-0cb654bc73e4}");
        /// <summary>
        /// Returns a GUID that represents the instant messaging addressof a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid IMAddress = new Guid("{4cbd96f7-09c6-4b5e-ad42-1cbe123de63a}");
        /// <summary>
        /// Returns a GUID that represents the home street address for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeAddressStreet = new Guid("{8c66e340-0985-4d68-af03-3050ece4862b}");
        /// <summary>
        /// Returns a GUID that represents the home city of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeAddressCity = new Guid("{5aeabc56-57c6-4861-bc12-bd72c30fc6bd}");
        /// <summary>
        /// Returns a GUID that represents the home state or province of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeAddressStateOrProvince = new Guid("{f5b36006-69b0-418c-bd4a-f25ca7e096bb}");
        /// <summary>
        /// Returns a GUID that represents information about the home postal code of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeAddressPostalCode = new Guid("{c0e4b4c6-6245-4846-8561-b8c6c01fefc1}");
        /// <summary>
        /// Returns a GUID that represents the home country of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid HomeAddressCountry = new Guid("{897ecfd7-4293-4782-b463-bd68440a5fed}");
        /// <summary>
        /// Returns a GUID that represents an alternative street address of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherAddressStreet = new Guid("{dff5dfc2-e2b7-4a19-bde7-76dabc90a3d2}");
        /// <summary>
        /// Returns a GUID that represents an alternative city address of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherAddressCity = new Guid("{90fa9a8e-aac0-4828-9cb4-78f98416affa}");
        /// <summary>
        /// Returns a GUID that represents the alternative state or province of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherAddressStateOrProvince = new Guid("{f45883bc-8733-4b77-ab5d-43613986aa12}");
        /// <summary>
        /// Returns a GUID that represents the alternative postal code of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherAddressPostalCode = new Guid("{0557c3f8-60c4-4dfb-b5ba-bf3c4e4386b1}");
        /// <summary>
        /// Returns a GUID that represents an alternative country of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OtherAddressCountry = new Guid("{3c0e9e00-8fcc-479f-9d8d-3447cda34c5b}");
        /// <summary>
        /// Returns a GUID that represents a second e-mail address for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Email2 = new Guid("{e232d6c8-9f49-4be2-bb28-b90570bcf167}");
        /// <summary>
        /// Returns a GUID that represents a third e-mail address for a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Email3 = new Guid("{8bd27dbd-29a0-4ccd-bcb4-03fe70c538b1}");
        /// <summary>
        /// Returns a GUID that represents the department name or identifier (ID) of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid ol_Department = new Guid("{c814b2cf-84c6-4f56-b4a4-c766938a97c5}");
        /// <summary>
        /// Returns a GUID that represents the identifier of the physical office of a person who is represented by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid Office = new Guid("{26169ab2-4bd2-4870-b077-10f49c8a5822}");
        /// <summary>
        /// Returns a GUID that represents the profession of a person who is referenced in a Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Profession = new Guid("{f0753a13-44b1-4269-82af-5c34c57b0c67}");
        /// <summary>
        /// Returns a GUID that represents the manager's name in the corporate hierarchy for a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid ManagersName = new Guid("{ba934502-d68d-4960-a54b-51e15fef5fd3}");
        /// <summary>
        /// Returns a GUID that represents the name of the assistant to a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid AssistantsName = new Guid("{2aea194d-e399-4f05-95af-94f87b1f2687}");
        /// <summary>
        /// Returns a GUID that represents the informal name of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Nickname = new Guid("{6b0a2cd7-a7f9-41ca-b932-f3bebb603793}");
        /// <summary>
        /// Returns a GUID that represents the name of the spouse of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid SpouseName = new Guid("{f590b1de-8e28-4c17-91bc-bf4096024b7e}");
        /// <summary>
        /// Returns a GUID that represents the birth date that is associated with a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Birthday = new Guid("{c4c7d925-bc1b-4f37-826d-ac49b4fb1bc1}");
        /// <summary>
        /// Returns a GUID that represents the corporate start date for the specified Windows SharePoint Services user.
        /// </summary>
        public static readonly Guid Anniversary = new Guid("{9d76802c-13c4-484a-9872-d7f9641c4672}");
        /// <summary>
        /// Returns a GUID that represents the gender of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Gender = new Guid("{23550288-91b5-4e7f-81f9-1a92661c4838}");
        /// <summary>
        /// Returns a GUID that represents the initials that are associated with a person who is represented by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Initials = new Guid("{7a282f86-69d9-40ff-ae1c-c746cf21256b}");
        /// <summary>
        /// Returns a GUID that represents information about the personal activities of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid Hobbies = new Guid("{203fa378-6eb8-4ed9-a4f9-221a4c1fbf46}");
        /// <summary>
        /// Returns a GUID that represents a field that contains the names of children who in turn are associated with a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid ChildrensNames = new Guid("{6440b402-8ec5-4d7a-83f4-afccb556b5cc}");
        /// <summary>
        /// Returns a GUID that represents the customized information in a field named UserField1 for a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid UserField1 = new Guid("{566656f5-17b3-4291-98a5-5074aadf77b3}");
        /// <summary>
        /// Returns a GUID that represents the customized information in a field named UserField2 for a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid UserField2 = new Guid("{182d1b9e-1718-4e11-b279-38f7ed0a20d6}");
        /// <summary>
        /// Returns a GUID that represents the customized information in a field named UserField3 for a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid UserField3 = new Guid("{a03eb53e-f123-4af9-9355-f92bd75c00b3}");
        /// <summary>
        /// Returns a GUID that represents the customized information in a field named UserField4 for a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid UserField4 = new Guid("{adefa4ca-14c3-4694-b531-f51b706efe9d}");
        /// <summary>
        /// Returns a GUID that represents the Government Identification number of a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid GovernmentIDNumber = new Guid("{da31d3c9-f9da-4c35-88d4-60aafa4c3f19}");
        /// <summary>
        /// Returns a GUID that represents the name of a computer network for a person who is referenced in a specified  Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid ComputerNetworkName = new Guid("{86a78395-c8ad-429e-abff-be09417b523e}");
        /// <summary>
        /// Returns a GUID that represents the name of the person who provided a referral for a person who is referenced by a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid ReferredBy = new Guid("{9b4cc5a9-1119-43e4-b2a8-412c4031f92b}");
        /// <summary>
        /// Returns a GUID that represents the employee identification number or organizational identification number that applies to a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid OrganizationalIDNumber = new Guid("{0850ae15-19dd-431f-9c2f-3aff3ae292ce}");
        public static readonly Guid CustomerID = new Guid("{81368791-7cbc-4230-981a-a7669ade9801}");
        /// <summary>
        /// Returns a GUID that represents the URL for the personal Web site of a person who is represented by a specified Windows SharePoint Services object.
        /// </summary>
        public static readonly Guid PersonalWebsite = new Guid("{5aa071d9-3254-40fb-82df-5cedeff0c41e}");
        /// <summary>
        /// Returns a GUID that represents a File Transfer Protocol (FTP) URL that is associated with a person who is referenced in a specified Windows SharePoint Services contact object.
        /// </summary>
        public static readonly Guid FTPSite = new Guid("{d733736e-4204-4812-9565-191567b27e33}");
        public static readonly Guid ParentVersionString = new Guid("{bc1a8efb-0f4c-49f8-a38f-7fe22af3d3e0}");
        public static readonly Guid ParentLeafName = new Guid("{774eab3a-855f-4a34-99da-69dc21043bec}");
        public static readonly Guid _DCDateCreated = new Guid("{9f8b4ee0-84b7-42c6-a094-5cbde2115eb9}");
        public static readonly Guid _Identifier = new Guid("{3c76805f-ad45-483a-9c85-7ac24506ce1a}");
        public static readonly Guid _Version = new Guid("{78be84b9-d70c-447b-8275-8dcd768b6f92}");
        public static readonly Guid _Revision = new Guid("{16b4ab96-0ce5-4c82-a836-f3117e8996ff}");
        public static readonly Guid _DCDateModified = new Guid("{810dbd02-bbf5-4c67-b1ce-5ad7c5a512b2}");
        public static readonly Guid _LastPrinted = new Guid("{b835f7c6-88a0-45d5-80c9-7ab4b2888b2b}");
        public static readonly Guid _Contributor = new Guid("{370b7779-0344-4b9f-8f2d-dc1c62eae801}");
        public static readonly Guid _Coverage = new Guid("{3b1d59c0-26b1-4de6-abbd-3edb4e2c6eca}");
        public static readonly Guid _Format = new Guid("{36111fdd-2c65-41ac-b7ef-48b9b8da4526}");
        public static readonly Guid _Publisher = new Guid("{2eedd0ae-4281-4b77-99be-68f8b3ad8a7a}");
        public static readonly Guid _Relation = new Guid("{5e75c854-6e9d-405d-b6c1-f8725bae5822}");
        public static readonly Guid _RightsManagement = new Guid("{ada3f0cb-6f95-4588-bb08-d97cc0623522}");
        public static readonly Guid _Source = new Guid("{b0a3c1db-faf1-48f0-9be1-47d2fc8cb5d6}");
        public static readonly Guid _ResourceType = new Guid("{edecec70-f6e2-4c3c-a4c7-f61a515dfaa9}");
        public static readonly Guid _EditMenuTableStart2 = new Guid("{1344423c-c7f9-4134-88e4-ad842e2d723c}");
        public static readonly Guid MyEditor = new Guid("{078b9dba-eb8c-4ec5-bfdd-8d220a3fcc5d}");
        public static readonly Guid ThumbnailExists = new Guid("{1f43cd21-53c5-44c5-8675-b8bb86083244}");
        public static readonly Guid AlternateThumbnailUrl = new Guid("{f39d44af-d3f3-4ae6-b43f-ac7330b5e9bd}");
        public static readonly Guid PreviewExists = new Guid("{3ca8efcd-96e8-414f-ba90-4c8c4a8bfef8}");
        public static readonly Guid IconOverlay = new Guid("{b77cdbcf-5dce-4937-85a7-9fc202705c91}");
        public static readonly Guid UIVersion = new Guid("{8e334549-c2bd-4110-9f61-672971be6504}");
        public static readonly Guid SortBehavior = new Guid("{423874f8-c300-4bfb-b7a1-42e2159e3b19}");
        public static readonly Guid FolderChildCount = new Guid("{960ff01f-2b6d-4f1b-9c3f-e19ad8927341}");
        public static readonly Guid ItemChildCount = new Guid("{b824e17e-a1b3-426e-aecf-f0184d900485}");
        public static readonly Guid EmailHeaders = new Guid("{e6985df4-cf66-4313-bcda-d89744d3b02f}");
        public static readonly Guid Predecessors = new Guid("{c3a92d97-2b77-4a25-9698-3ab54874bc6f}");
        public static readonly Guid MobilePhone = new Guid("{bf03d3ca-aa6e-4845-809a-b4378b37ce08}");
        public static readonly Guid wic_System_Copyright = new Guid("{f08ab41d-9a03-49ae-9413-6cd284a15625}");
        public static readonly Guid PreviewOnForm = new Guid("{8c0d0aac-9b76-4951-927a-2490abe13c0b}");
        public static readonly Guid ThumbnailOnForm = new Guid("{9941082a-4160-46a1-a5b2-03394bfdf7ee}");
        public static readonly Guid NoCodeVisibility = new Guid("{a05a8639-088a-4aea-b8a9-afc888971c81}");
        public static readonly Guid AssociatedListId = new Guid("{b75067a2-e23b-499f-aa07-4ceb6c79e0b3}");
        public static readonly Guid RestrictContentTypeId = new Guid("{8b02a33c-accd-4b73-bcae-6932c7aab812}");
        public static readonly Guid WorkflowDisplayName = new Guid("{5263cd09-a770-4549-b012-d9f3df3d8df6}");
        public static readonly Guid ParticipantsPicker = new Guid("{8137f7ad-9170-4c1d-a17b-4ca7f557bc88}");
        public static readonly Guid Participants = new Guid("{453c2d71-c41e-46bc-97c1-a5a9535053a3}");
        public static readonly Guid Facilities = new Guid("{a4e7b3e1-1b0a-4ffa-8426-c94d4cb8cc57}");
        public static readonly Guid FreeBusy = new Guid("{393003f9-6ccb-4ea9-9623-704aa4748dec}");
        public static readonly Guid Overbook = new Guid("{d8cd5bcf-3768-4d6c-a8aa-fefa3c793d8d}");
        public static readonly Guid GbwLocation = new Guid("{afaa4198-9797-4e45-9825-8f7e7b0f5dd5}");
        public static readonly Guid GbwCategory = new Guid("{7fc04acf-6b4f-418c-8dc5-ecfb0085bb51}");
        public static readonly Guid WhatsNew = new Guid("{cf68a174-123b-413e-9ec1-b43e3a3175d7}");
        public static readonly Guid DueDate = new Guid("{c1e86ea6-7603-493c-ab5d-db4bbfe8f96a}");
        public static readonly Guid Confidential = new Guid("{9b0e6471-c5c5-42ef-9ade-63170bf28819}");
        public static readonly Guid AllowEditing = new Guid("{7266b59c-030b-4ca3-bc09-bb8e76ad969b}");
        public static readonly Guid V4SendTo = new Guid("{e0f298a5-7e3e-4895-9ff8-90d88ec4526d}");
        public static readonly Guid Confirmations = new Guid("{ef7465d3-5d54-487b-b081-ade80acae88e}");
        public static readonly Guid V4CallTo = new Guid("{7111aa1b-e7ae-4b69-acaf-db669b76e03a}");
        public static readonly Guid ConfirmedTo = new Guid("{1b89212c-1c67-487a-8c14-4d30bf4ef223}");
        public static readonly Guid CallBack = new Guid("{274b7e21-284a-4c49-bec6-f1f2cb6fc344}");
        public static readonly Guid Detail = new Guid("{6529a881-d745-4117-a552-3dcc7110e9b8}");
        public static readonly Guid CallTime = new Guid("{63fc6806-db53-4d0d-b18b-eaf90e96ddf5}");
        public static readonly Guid Resolved = new Guid("{a6fd2bb9-c701-4168-99cc-242e42f7671a}");
        public static readonly Guid ResolvedBy = new Guid("{b4fa187b-eb65-478e-8bc6-93b0da320f03}");
        public static readonly Guid ResolvedDate = new Guid("{c4995c71-4c5c-4e9f-afc1-a9033f2bfde5}");
        public static readonly Guid Description = new Guid("{3f155110-a6a2-4d70-926c-94648101f0e8}");
        public static readonly Guid HolidayDate = new Guid("{335e22c3-b8a4-4234-9790-7a03eeb7b0d4}");
        public static readonly Guid V4HolidayDate = new Guid("{492b1ac0-c594-4013-a2b6-ea70f5a8a506}");
        public static readonly Guid IsNonWorkingDay = new Guid("{baf7091c-01fb-4831-a975-08254f87f234}");
        public static readonly Guid UserName = new Guid("{211a8cfc-93b7-4173-9254-0bfe2d1643da}");
        public static readonly Guid Date = new Guid("{2139e5cc-6c75-4a65-b84c-00fe93027db3}");
        public static readonly Guid DayOfWeek = new Guid("{61fc45dd-b33d-4679-8646-be9e6584fadd}");
        public static readonly Guid Start = new Guid("{05e6336c-d22e-478e-9414-366762883b3f}");
        public static readonly Guid End = new Guid("{04b29608-b1e8-4ff9-90d5-5328096dd5ac}");
        public static readonly Guid In = new Guid("{ee394fd4-4c11-4d8e-baff-83270c1921aa}");
        public static readonly Guid Out = new Guid("{fde05b9b-52bf-43dc-9b96-bb35fa7aa05d}");
        public static readonly Guid Break = new Guid("{9b12fb06-254e-43b3-bfc8-8eea422ebc9f}");
        public static readonly Guid ScheduledWork = new Guid("{3bdf7bd3-f229-419e-8e12-3dfecb49ed38}");
        public static readonly Guid Overtime = new Guid("{35d79e8b-3701-4659-9c27-c070ed3c2bfa}");
        public static readonly Guid NightWork = new Guid("{aaa68c08-6276-4337-9bce-b9cd852c7328}");
        public static readonly Guid HolidayWork = new Guid("{b5a7350f-2716-46ca-9c42-66bb39d042ec}");
        public static readonly Guid HolidayNightWork = new Guid("{dc9100ec-251d-4e81-a6cb-d967a065ba24}");
        public static readonly Guid Late = new Guid("{df7f27a4-d87b-4a97-947b-13d1d4f7e6de}");
        public static readonly Guid LeaveEarly = new Guid("{a2a86efe-c28e-4dde-ab56-0afa31664bbc}");
        public static readonly Guid Oof = new Guid("{63c1c608-df6f-4cfa-bcab-fdbf9c223e31}");
        public static readonly Guid Vacation = new Guid("{dfd58778-bf8e-4769-8265-09ac03159eed}");
        public static readonly Guid NumberOfVacation = new Guid("{44e16d52-da1b-4e72-8bdb-89a3b77ec8b0}");
        public static readonly Guid ShortComment = new Guid("{691b9a4b-512e-4341-b3f1-68914130d5b2}");
        public static readonly Guid ListType = new Guid("{81dde544-1e25-4765-b5fd-ba613198d850}");
        public static readonly Guid Content = new Guid("{7650d41a-fa26-4c72-a641-af4e93dc7053}");
        public static readonly Guid MobileContent = new Guid("{53a2a512-d395-4852-8714-d4c27e7585f3}");
        public static readonly Guid Whereabout = new Guid("{e2a07293-596a-4c59-9089-5c4f9339077f}");
        public static readonly Guid From = new Guid("{4cd541b9-c8ee-468f-bee6-33f3b9baa722}");
        public static readonly Guid GoFromHome = new Guid("{6570d35e-7f0a-4123-93c9-f53ffa5810d3}");
        public static readonly Guid Until = new Guid("{fe3344ab-b468-471f-8fa5-9b506c7d1557}");
        public static readonly Guid GoingHome = new Guid("{2ead592e-f05c-41a2-9817-e06dac25bc19}");
        public static readonly Guid ContactInfo = new Guid("{e1a85174-b8d0-4962-9ce6-758f8b612725}");
        public static readonly Guid IMEDisplay = new Guid("{90244050-709c-4837-9316-93863fbd3da6}");
        public static readonly Guid IMEComment1 = new Guid("{d2433b20-3f02-4432-817d-369f104a2dcd}");
        public static readonly Guid IMEComment2 = new Guid("{e2c93917-cf32-4b29-be5c-d71f1bac7714}");
        public static readonly Guid IMEComment3 = new Guid("{7c52f61a-e1e0-4341-9e2f-9b36cddfdd7c}");
        public static readonly Guid IMEUrl = new Guid("{84b0fe85-6b16-40c3-8507-e56c5bbc482e}");
        public static readonly Guid IMEPos = new Guid("{f3cdbcfd-f456-45f4-9000-b6f34bb95d84}");
        public static readonly Guid HealthRuleService = new Guid("{2d6e61d0-be31-460c-ab8b-77d8b369f517}");
        public static readonly Guid HealthRuleType = new Guid("{7dd0a092-8704-4ed2-8253-ac309150ac59}");
        public static readonly Guid HealthRuleScope = new Guid("{e59f08c9-fa34-4f94-a00a-f6458b1d3c56}");
        public static readonly Guid HealthRuleSchedule = new Guid("{26761ba3-729d-4bfc-9658-77b55e01f8d5}");
        public static readonly Guid HealthReportServers = new Guid("{84a318aa-9035-4529-98b9-e08bb20a5da0}");
        public static readonly Guid HealthReportServices = new Guid("{e2b0b450-6795-4b86-86b7-3c21ab1797fb}");
        public static readonly Guid HealthReportCategory = new Guid("{a63505f2-f42c-4d94-b03b-78ba2c73d40e}");
        public static readonly Guid HealthReportExplanation = new Guid("{b4c8faec-5d60-49ee-a5fb-6165f5c3e6a9}");
        public static readonly Guid HealthReportRemedy = new Guid("{8aa22caa-8000-44c9-b343-a7705bbed863}");
        public static readonly Guid HealthRuleReportLink = new Guid("{cf4ff575-f1f5-4c5b-b595-54bbcccd0c62}");
        public static readonly Guid HealthReportSeverityIcon = new Guid("{89efcbd9-9796-41f0-b569-65325f1882dc}");
        public static readonly Guid HealthReportSeverity = new Guid("{505423c5-f085-48b9-9432-12073d643ba5}");
        public static readonly Guid HealthRuleAutoRepairEnabled = new Guid("{1e41a55e-ef71-4740-b65a-d11e24c1d00d}");
        public static readonly Guid HealthRuleCheckEnabled = new Guid("{7b2b1712-a73d-4ad7-a9d0-662f0291713d}");
        public static readonly Guid HealthRuleVersion = new Guid("{6b6b1455-09ee-43b7-beea-4dc97456de2f}");
        public static readonly Guid XSLStyleCategory = new Guid("{dfffbbfb-0cc3-4ce7-8cb3-a2958fb726a1}");
        public static readonly Guid XSLStyleWPType = new Guid("{4499086f-9ac1-41df-86c3-d8c1f8fc769a}");
        public static readonly Guid XSLStyleIconUrl = new Guid("{3dfb3e11-9ccd-4404-b44a-a71f6399ea56}");
        public static readonly Guid XSLStyleBaseView = new Guid("{4630e6ac-e543-4667-935a-2cc665e9b755}");
        public static readonly Guid XSLStyleRequiredFields = new Guid("{acb9088a-a171-4b99-aa7a-10388586bc74}");
        public static readonly Guid ParentID = new Guid("{fd447db5-3908-4b47-8f8c-a5895ed0aa6a}");
        public static readonly Guid AppAuthor = new Guid("{6bfaba20-36bf-44b5-a1b2-eb6346d49716}");
        public static readonly Guid AppEditor = new Guid("{e08400f3-c779-4ed2-a18c-ab7f34caa318}");
        public static readonly Guid NoCrawl = new Guid("{b0e12a3b-cf63-47d1-8418-4ef850d87a3c}");
        public static readonly Guid PrincipalCount = new Guid("{dcc67ebd-247f-4bee-8626-85ff6f69fbb6}");
        public static readonly Guid Checkmark = new Guid("{ebf1c037-47eb-4355-998d-47ce9f2cc047}");
        public static readonly Guid RelatedLinks = new Guid("{1ad7c220-c893-4c15-b95c-b69b992bdee2}");
        public static readonly Guid MUILanguages = new Guid("{fb005daa-caf9-4ecd-84d5-6bdd2eb3dce7}");
        public static readonly Guid ContentLanguages = new Guid("{58073ebd-b204-4899-bc77-54402c61e9e9}");
        public static readonly Guid UserInfoHidden = new Guid("{e8a80787-5f99-459a-af8d-b830157ed45f}");
        public static readonly Guid IsFeatured = new Guid("{5a034ff8-d7a4-4d69-ab26-5f5a043b572d}");
        public static readonly Guid DisplayTemplateJSTemplateHidden = new Guid("{3d0684f7-ca97-413d-9d03-d00f480059ae}");
        public static readonly Guid DisplayTemplateJSTargetControlType = new Guid("{0e49b273-3102-4b7d-b609-2e05dd1a17d9}");
        public static readonly Guid DisplayTemplateJSIconUrl = new Guid("{57468ccb-0c02-422c-ba0a-61a44ba41784}");
        public static readonly Guid DisplayTemplateJSTemplateType = new Guid("{d63173ac-b914-4f90-9cf8-4ff4352e41a3}");
        public static readonly Guid DisplayTemplateJSTargetScope = new Guid("{df8bd7e5-b3db-4a94-afb4-7296397d829d}");
        public static readonly Guid DisplayTemplateJSTargetListTemplate = new Guid("{9f927425-78e9-49c3-b03b-65e1211394e1}");
        public static readonly Guid DisplayTemplateJSTargetContentType = new Guid("{ed095cf7-534e-460b-965f-f14269e70f5a}");
        public static readonly Guid DisplayTemplateJSConfigurationUrl = new Guid("{0f2f686a-3921-432e-85fd-9c535bf671b2}");
        public static readonly Guid DefaultCssFile = new Guid("{cc10b158-50b4-4f02-8f3a-b9b6c3102628}");
        public static readonly Guid RelatedItems = new Guid("{d2a04afc-9a05-48c8-a7fa-fa98f9496141}");
        public static readonly Guid PreviouslyAssignedTo = new Guid("{1982e408-0f94-4149-8349-16f301d89134}");
        public static readonly Guid SkipFeatureDeployment = new Guid("{eda244c9-5a10-4057-8b96-a118a3b31454");
        public static readonly Guid AppPackageErrorMessage = new Guid("{80ac5bc8-73dc-4967-950f-3f6ec32c19f4");
        public static readonly Guid AppPermissionXML = new Guid("{bad900c5-34c6-4d4b-b84a-698b5da331fe");
        public static readonly Guid AppSubtype = new Guid("{de81f63e-c919-4f47-80fb-95ef0ef506f7");
        public static readonly Guid AppSubtypeID = new Guid("{e026628d-a28f-4ad3-b34b-a95b3b009ccb");
        public static readonly Guid AppTitleInfo = new Guid("{d1807ced-4fae-49ab-8307-9ec769c8d838");
        public static readonly Guid AppVersion = new Guid("{8416ff0b-f1ab-49fc-a7ef-a744f871f058");
        public static readonly Guid AssetID = new Guid("{d3321fff-46f3-407d-89ef-fff1e075a018");
        public static readonly Guid IsAutoHostedApp = new Guid("{5a7ad4d7-13d0-48e6-864f-ad693b491788");
        public static readonly Guid AadApplicationAppId = new Guid("{23410d43-d091-46aa-a459-73c200fa2e26");
        public static readonly Guid AadApplicationId = new Guid("{33bda339-efe6-4a1b-acc8-6c79a1dbae98");
        public static readonly Guid AadServicePrincipalId = new Guid("{32434713-0003-4acc-9976-9ead1309888b");
        public static readonly Guid SharePointAppCategory = new Guid("{7cbaba9f-d082-4e55-96b9-62ca9a541769");
        public static readonly Guid IsClientSideSolution = new Guid("{2b6eea14-803b-47d2-959f-4947bf4e978c");
        public static readonly Guid ClientComponentDeveloperMpnId = new Guid("{28e60c99-8bf9-450d-b5aa-125953a4d4f1");
        public static readonly Guid ClientComponentDeveloperName = new Guid("{5dc8760e-c00e-4e80-ae4c-757906c24712");
        public static readonly Guid ClientComponentDeveloperPrivacyUrl = new Guid("{d2c76127-be32-47ed-a153-a45f55e74c99");
        public static readonly Guid ClientComponentDeveloperTermsOfUseUrl = new Guid("{124a6f7f-99f7-4ffc-b00e-27c7e7eb121f");
        public static readonly Guid ClientComponentDeveloperWebsiteUrl = new Guid("{2639fc06-3b49-4365-834b-149fd1955bd8");
        public static readonly Guid TenantWideExtensionComponentId = new Guid("{a146d12a-1902-41a9-920a-155738eff298");
        public static readonly Guid ClientComponentId = new Guid("{1c640fff-ceda-4798-bf7a-096d3b1b823e");
        public static readonly Guid ClientComponentManifest = new Guid("{f755dd56-4435-4f64-b4c5-df939a12540d");
        public static readonly Guid ClientComponentName = new Guid("{d2c5909d-d4b5-46b1-adee-13d4bc7394c9");
        public static readonly Guid TenantWideExtensionComponentProperties = new Guid("{ef74e985-ebf0-4c48-b722-28ec6266ffa4");
        public static readonly Guid ClientComponentType = new Guid("{1b9bece7-97b2-4a80-8282-e059fd8b73ae");
        public static readonly Guid ClientComponentConnectedTeamsAppId = new Guid("{296f2707-0bb2-488d-81d0-0af188aa5b7a");
        public static readonly Guid ContainsTeamsManifest = new Guid("{4618bc96-56d1-49e1-8385-4e9be7b4d604");
        public static readonly Guid ContainsTenantWideExtension = new Guid("{84316157-4f2b-42c4-b353-b593b357de01");
        public static readonly Guid ContainsVivaManifest = new Guid("{bf222d9b-8e25-4b89-8611-43c96bb550d3");
        public static readonly Guid IsClientSideSolutionCurrentVersionDeployed = new Guid("{96ccf93b-0eb3-4a81-b579-5f1bad70dbcb");
        public static readonly Guid IsDefaultAppMetadataLocale = new Guid("{2e055cad-648e-4723-aa3a-917dd9e809b3");
        public static readonly Guid IsClientSideSolutionDeployed = new Guid("{b2fb0f79-cb89-4a24-a048-561d4b3505f7");
        public static readonly Guid AppDescription = new Guid("{5b90ee3e-12a0-44c7-86c6-d8fbc461777d");
        public static readonly Guid TenantWideExtensionDisabled = new Guid("{41494365-37f4-4a7d-b197-d54650fa0f1d");
        public static readonly Guid Disabled = new Guid("{051adcd1-34e2-48ca-a10d-973ba0c58227");
        public static readonly Guid IsAppPackageEnabled = new Guid("{0985248f-8044-4c6f-b535-eb3ac2773d44");
        public static readonly Guid ExternalContentDomains = new Guid("{d1897cf8-37a7-486b-98b3-323672559684");
        public static readonly Guid FeatureDescription = new Guid("{c43f8f87-1ebc-4c97-b222-b33884e93053");
        public static readonly Guid FeatureId = new Guid("{b53f4070-967c-4db3-98df-3cc990e80c5e");
        public static readonly Guid FeatureName = new Guid("{95668939-6451-4ae1-91fa-c5923f9e0be7");
        public static readonly Guid FeatureScope = new Guid("{717f079d-2f2e-45af-a59c-6a7a9e817897");
        public static readonly Guid IsFeaturedApp = new Guid("{b9d82342-9580-4ba1-9742-174cfccbc77f");
        public static readonly Guid IsClientSideSolutionGeneratedCardDesign = new Guid("{c6fdc819-8901-4abf-abd6-e0764b45fd45");
        public static readonly Guid ClientComponentIsGeneratedCardDesign = new Guid("{63d502fb-1cb0-4312-9ed5-56cc7b6509ad");
        public static readonly Guid TenantWideExtensionHostProperties = new Guid("{4740858a-940a-48da-8553-2e10fec5a056");
        public static readonly Guid AppThumbnailURL = new Guid("{22fa2e37-2293-4e0b-90e9-fa00ccbd85b6");
        public static readonly Guid AppImageURL1 = new Guid("{e44b4501-af7d-4b30-a8a2-e70b59dae7bf");
        public static readonly Guid AppImageURL2 = new Guid("{047d5b20-5e82-4ea1-9841-f62c9656c41a");
        public static readonly Guid AppImageURL3 = new Guid("{873fd80b-3f9e-4ada-88f4-3ee3c15522ac");
        public static readonly Guid AppImageURL4 = new Guid("{307ecf23-25d5-4d1d-89f5-693eb977d6f5");
        public static readonly Guid AppImageURL5 = new Guid("{1fc17524-afa8-4e6f-ba64-a77dd42f310b");
        public static readonly Guid IsolatedDomain = new Guid("{68021728-c04c-45bb-b51f-7bcb5ab997b1");
        public static readonly Guid TenantWideExtensionListTemplate = new Guid("{8ae577f5-25a2-4122-8705-3460142b2069");
        public static readonly Guid TenantWideExtensionLocation = new Guid("{6b75393a-f755-4290-bfda-702ad4fdeac3");
        public static readonly Guid AppMetadataLocale = new Guid("{14c6cd06-7417-42c1-a051-89e455fd1090");
        public static readonly Guid AppCatalog_IsNewVersionAvailable = new Guid("{bb941810-6999-4132-915d-b54a46a56233");
        public static readonly Guid PackageDefaultSkipFeatureDeployment = new Guid("{a39bc357-357e-403a-8146-289f121b9040");
        public static readonly Guid AppPackageHash = new Guid("{0a95ebcf-b06d-479c-8196-c5783c96f751");
        public static readonly Guid AppPrerequisitesXML = new Guid("{83cad291-c898-4d17-9208-d6786fc17145");
        public static readonly Guid PreviouslySetClientComponentSkipFeatureDeployment = new Guid("{d3cd6df8-cc3b-4dff-be20-8ff488a52955");
        public static readonly Guid AppProductID = new Guid("{031cd33c-6486-4b08-af1b-8d2b05eb3854");
        public static readonly Guid IsProviderHostedClientSideSolution = new Guid("{886b911a-e83c-4a3b-8931-25c832bacac6");
        public static readonly Guid AppPublisher = new Guid("{b1402c55-9750-4346-b6c3-120b2da47b73");
        public static readonly Guid ReturnIfCustomScriptDisabled = new Guid("{274a482a-ac49-4db7-bbb5-8bfbdd27b45b");
        public static readonly Guid TenantWideExtensionSequence = new Guid("{89aa6026-7b00-4d24-8b4a-7c7de7880373");
        public static readonly Guid AppShortDescription = new Guid("{cf57b756-a2b6-43bd-ace3-6fd4d8ad90f8");
        public static readonly Guid ClientComponentSkipFeatureDeployment = new Guid("{cd58f851-9090-49ed-a54e-fe488f662d4a");
        public static readonly Guid SolutionId = new Guid("{fd34ca2f-4e81-40fa-b16f-4c455235da19");
        public static readonly Guid SolutionName = new Guid("{f3d82830-df08-4612-b895-61198f55c723");
        public static readonly Guid SolutionVersion = new Guid("{f92c6a2a-1749-428e-b7da-3e44c3beaa8f");
        public static readonly Guid StoreAssetId = new Guid("{3aa8efc0-1a31-42bc-9755-865f1a22a3f5");
        public static readonly Guid AppSupportURL = new Guid("{cd5e6c48-2138-4bf7-bfc3-4e5e9ece7e1c");
        public static readonly Guid AppSupportedLocales = new Guid("{dc652701-23cb-44ee-bf83-9e740746cd4c");
        public static readonly Guid ComponentSupportsFullPage = new Guid("{cbc0003b-9d7e-4c33-bea9-a37effdb7fb1");
        public static readonly Guid ComponentSupportsTeamsTab = new Guid("{f689d549-103c-499c-96ed-709615986111");
        public static readonly Guid SupportsTeamsTabs = new Guid("{12592a6a-af80-430a-baea-657bdcec791d");
        public static readonly Guid ComponentSupportsWebPart = new Guid("{c8f457e4-a928-45a0-9fda-496bdf220dc2");
        public static readonly Guid UniqueSolutionId = new Guid("{55c702d1-fd66-4d3d-9e78-a0d70bab963a");
        public static readonly Guid IsValidAppPackage = new Guid("{01df698e-7495-428c-a04f-99c8eda6b3cd");
        public static readonly Guid AppVideoURL = new Guid("{831a8efd-9456-43fa-b584-1cbdd4292519");
        public static readonly Guid WebApiPermissionScopesNote = new Guid("{5d88909c-873a-4262-bc81-9b7c2ad04105");
        public static readonly Guid TenantWideExtensionWebTemplate = new Guid("{8b97c8ba-e0c8-4c35-9b65-efee64f1b610");
        public static readonly Guid WebApiPermissionRequests = new Guid("{f8a04db8-4b6e-4613-9232-02811628a7f6");
        private static Dictionary<Guid, bool> s_dict = null;

        /// <summary>
        /// This method returns a Boolean value that specifies whether or not the current object matches the specified GUID. This value is used as a file identifier for an object that is associated with a Windows SharePoint Services Web site.
        /// </summary>
        /// 
        /// <returns>
        /// Returns a GUID.
        /// </returns>
        /// <param name="fid">File identifier.</param>
        public static bool Contains(Guid fid)
        {
            if (s_dict == null)
            {
                s_dict = new Dictionary<Guid, bool>(523)
                {
                    [DisplayTemplateJSTargetListTemplate] = true,
                    [Editor] = true,
                    [WebPage] = true,
                    [Profession] = true,
                    [IsNonWorkingDay] = true,
                    [CallTime] = true,
                    [ImageHeight] = true,
                    [EndDate] = true,
                    [Modified_x0020_By] = true,
                    [Last_x0020_Modified] = true,
                    [ThumbnailExists] = true,
                    [RelevantMessages] = true,
                    [ContentLanguages] = true,
                    [MiddleName] = true,
                    [HolidayWork] = true,
                    [AllowEditing] = true,
                    [HealthReportSeverity] = true,
                    [_EditMenuTableEnd] = true,
                    [OffsiteParticipant] = true,
                    [CallBack] = true,
                    [Location] = true,
                    [Comments] = true,
                    [ParentID] = true,
                    [OtherAddressCity] = true,
                    [LinkIssueIDNoMenu] = true,
                    [Created_x0020_Date] = true,
                    [Gender] = true,
                    [WorkflowDisplayName] = true,
                    [SpouseName] = true,
                    [Service] = true,
                    [Date] = true,
                    [HTML_x0020_File_x0020_Type] = true,
                    [Resolved] = true,
                    [User] = true,
                    [RelatedItems] = true,
                    [URL] = true,
                    [Detail] = true,
                    [RecurrenceID] = true,
                    [AppAuthor] = true,
                    [HealthRuleSchedule] = true,
                    [ParentItemEditor] = true,
                    [DLC_Duration] = true,
                    [HomeAddressStateOrProvince] = true,
                    [Company] = true,
                    [Until] = true,
                    [CheckoutUser] = true,
                    [ThreadingControls] = true,
                    [FirstName] = true,
                    [From] = true,
                    [DefaultCssFile] = true,
                    [DiscussionTitle] = true,
                    [FullBody] = true,
                    [WorkflowVersion] = true,
                    [VirusStatus] = true,
                    [FirstNamePhonetic] = true,
                    [DisplayTemplateJSIconUrl] = true,
                    [End] = true,
                    [EncodedAbsThumbnailUrl] = true,
                    [Description] = true,
                    [DisplayTemplateJSTargetContentType] = true,
                    [V4HolidayDate] = true,
                    [EmailSubject] = true,
                    [IMEComment1] = true,
                    [ThreadTopic] = true,
                    [List] = true,
                    [Oof] = true,
                    [ContactInfo] = true,
                    [SendEmailNotification] = true,
                    [_HasCopyDestinations] = true,
                    [ParentFolderId] = true,
                    [NoCodeVisibility] = true,
                    [AttendeeStatus] = true,
                    [PercentComplete] = true,
                    [Body] = true,
                    [HealthReportCategory] = true,
                    [_CheckinComment] = true,
                    [_Revision] = true,
                    [Expires] = true,
                    [Email2] = true,
                    [HomeAddressCity] = true,
                    [Whereabout] = true,
                    [ComputerNetworkName] = true,
                    [File_x0020_Type] = true,
                    [Out] = true,
                    [AdminTaskDescription] = true,
                    [RelatedIssues] = true,
                    [DisplayTemplateJSConfigurationUrl] = true,
                    [_ModerationStatus] = true,
                    [DisplayTemplateJSTargetScope] = true,
                    [ParentItemID] = true,
                    [WorkflowItemId] = true,
                    [ShortestThreadIndexIdLookup] = true,
                    [Workspace] = true,
                    [OrganizationalIDNumber] = true,
                    [ScheduledWork] = true,
                    [Role] = true,
                    [MobilePhone] = true,
                    [Break] = true,
                    [IMEComment3] = true,
                    [RadioNumber] = true,
                    [SipAddress] = true,
                    [_Comments] = true,
                    [GoFromHome] = true,
                    [HealthRuleReportLink] = true,
                    [ReferredBy] = true,
                    [GoingHome] = true,
                    [WorkState] = true,
                    [ImageWidth] = true,
                    [ShortestThreadIndexId] = true,
                    [UserField4] = true,
                    [_Publisher] = true,
                    [ThreadIndex] = true,
                    [WorkflowOutcome] = true,
                    [AssignedTo] = true,
                    [SelectedFlag] = true,
                    [Keywords] = true,
                    [SelectTitle] = true,
                    [HomeAddressStreet] = true,
                    [ID] = true,
                    [Thumbnail] = true,
                    [TaskCompanies] = true,
                    [LastReplyBy] = true,
                    [IMEComment2] = true,
                    [ConnectionType] = true,
                    [UserField3] = true,
                    [BaseAssociationGuid] = true,
                    [MyEditor] = true,
                    [V4SendTo] = true,
                    [HasCustomEmailBody] = true,
                    [WorkflowName] = true,
                    [GbwCategory] = true,
                    [MessageId] = true,
                    [PreviewOnForm] = true,
                    [Indentation] = true,
                    [OtherAddressCountry] = true,
                    [EmailBody] = true,
                    [_Coverage] = true,
                    [fAllDayEvent] = true,
                    [PendingModTime] = true,
                    [BillingInformation] = true,
                    [Combine] = true,
                    [URLwMenu] = true,
                    [FullName] = true,
                    [OtherAddressPostalCode] = true,
                    [LinkFilename] = true,
                    [HomeAddressCountry] = true,
                    [_EditMenuTableStart] = true,
                    [_CopySource] = true,
                    [Author] = true,
                    [EmailReferences] = true,
                    [Department] = true,
                    [HealthRuleVersion] = true,
                    [CustomerID] = true,
                    [Modified] = true,
                    [Priority] = true,
                    [RulesUrl] = true,
                    [_Author] = true,
                    [AdminTaskAction] = true,
                    [PersonViewMinimal] = true,
                    [HealthRuleAutoRepairEnabled] = true,
                    [LinkDiscussionTitleNoMenu] = true,
                    [Home2Number] = true,
                    [GovernmentIDNumber] = true,
                    [Confirmations] = true,
                    [WorkflowTemplate] = true,
                    [XSLStyleIconUrl] = true,
                    [PublishedDate] = true,
                    [OtherFaxNumber] = true,
                    [PrincipalCount] = true,
                    [ParentLeafName] = true,
                    [DisplayTemplateJSTargetControlType] = true,
                    [XSLStyleBaseView] = true,
                    [_Format] = true,
                    [NameOrTitle] = true,
                    [LeaveEarly] = true,
                    [WorkflowInstance] = true,
                    [_SharedFileIndex] = true,
                    [PagerNumber] = true,
                    [EncodedAbsWebImgUrl] = true,
                    [Participants] = true,
                    [RepairDocument] = true,
                    [HealthReportExplanation] = true,
                    [ContentType] = true,
                    [_RightsManagement] = true,
                    [LinkDiscussionTitle2] = true,
                    [Purpose] = true,
                    [_LastPrinted] = true,
                    [PersonalWebsite] = true,
                    [ConfirmedTo] = true,
                    [Group] = true,
                    [TaskDueDate] = true,
                    [ShowCombineView] = true,
                    [LinkTitleNoMenu] = true,
                    [FileDirRef] = true,
                    [Name] = true,
                    [TaskType] = true,
                    [FileLeafRef] = true,
                    [TemplateUrl] = true,
                    [Overtime] = true,
                    [AlternateThumbnailUrl] = true,
                    [CallbackNumber] = true,
                    [Hobbies] = true,
                    [ShortComment] = true,
                    [_EditMenuTableStart2] = true,
                    [_UIVersionString] = true,
                    [WorkflowInstanceID] = true,
                    [XMLTZone] = true,
                    [EmailCalendarSequence] = true,
                    [wic_System_Copyright] = true,
                    [Confidential] = true,
                    [WorkflowLink] = true,
                    [ResolvedDate] = true,
                    [WorkZip] = true,
                    [EmailTo] = true,
                    [Suffix] = true,
                    [LastNamePhonetic] = true,
                    [Category] = true,
                    [V3Comments] = true,
                    [Mileage] = true,
                    [Deleted] = true,
                    [SortBehavior] = true,
                    [WorkFax] = true,
                    [_Relation] = true,
                    [CellPhone] = true,
                    [WorkspaceLink] = true,
                    [ol_Department] = true,
                    [In] = true,
                    [EmailFrom] = true,
                    [Office] = true,
                    [CompanyNumber] = true,
                    [Facilities] = true,
                    [HolidayNightWork] = true,
                    [DiscussionTitleLookup] = true,
                    [FTPSite] = true,
                    [WorkCity] = true,
                    [XomlUrl] = true,
                    [ContentTypeId] = true,
                    [UniqueId] = true,
                    [StatusBar] = true,
                    [EmailCalendarUid] = true,
                    [Vacation] = true,
                    [FreeBusy] = true,
                    [_Photo] = true,
                    [Comment] = true,
                    [Overbook] = true,
                    [NoCrawl] = true,
                    [HealthRuleScope] = true,
                    [TimeZone] = true,
                    [ISDNNumber] = true,
                    [RecurrenceData] = true,
                    [EMail] = true,
                    [_IsCurrentVersion] = true,
                    [File_x0020_Size] = true,
                    [WorkCountry] = true,
                    [NightWork] = true,
                    [AssociatedListId] = true,
                    [owshiddenversion] = true,
                    [AdminTaskOrder] = true,
                    [IsAnswered] = true,
                    [LinkFilenameNoMenu] = true,
                    [DueDate] = true,
                    [Start] = true,
                    [OtherAddressStateOrProvince] = true,
                    [ChildrensNames] = true,
                    [OtherAddressStreet] = true,
                    [ScopeId] = true,
                    [IconOverlay] = true,
                    [Threading] = true,
                    [_DCDateCreated] = true,
                    [JobTitle] = true,
                    [TaskStatus] = true,
                    [Outcome] = true,
                    [AssistantsName] = true,
                    [MessageBody] = true,
                    [Initials] = true,
                    [IsSiteAdmin] = true,
                    [PermMask] = true,
                    [RestrictContentTypeId] = true,
                    [Data] = true,
                    [BodyAndMore] = true,
                    [_Level] = true,
                    [ExtendedProperties] = true,
                    [IsQuestion] = true,
                    [EmailHeaders] = true,
                    [UIVersion] = true,
                    [_Version] = true,
                    [WorkflowAssociation] = true,
                    [_Contributor] = true,
                    [CompanyPhonetic] = true,
                    [ResolvedBy] = true,
                    [DecisionStatus] = true,
                    [Item] = true,
                    [ServerUrl] = true,
                    [AssistantNumber] = true,
                    [_UIVersion] = true,
                    [EventCanceled] = true,
                    [UID] = true,
                    [ReplyNoGif] = true,
                    [IsFeatured] = true,
                    [BaseName] = true,
                    [EmailSender] = true,
                    [Event] = true,
                    [ParticipantsPicker] = true,
                    [IndentLevel] = true,
                    [ActualWork] = true,
                    [V4CallTo] = true,
                    [Occurred] = true,
                    [EmailCc] = true,
                    [ToggleQuotedText] = true,
                    [LinkDiscussionTitle] = true,
                    [Title] = true,
                    [CarNumber] = true,
                    [UserField2] = true,
                    [fRecurrence] = true,
                    [IssueStatus] = true,
                    [ShowRepairView] = true,
                    [XSLStyleCategory] = true,
                    [BestAnswerId] = true,
                    [Subject] = true,
                    [Email3] = true,
                    [Anniversary] = true,
                    [Order] = true,
                    [HealthRuleService] = true,
                    [TrimmedBody] = true,
                    [_Category] = true,
                    [FileRef] = true,
                    [LimitedBody] = true,
                    [ManagersName] = true,
                    [_Status] = true,
                    [MasterSeriesItemID] = true,
                    [WorkflowListId] = true,
                    [Picture] = true,
                    [FormURN] = true,
                    [TTYTDDNumber] = true,
                    [OtherNumber] = true,
                    [Attachments] = true,
                    [URLNoMenu] = true,
                    [HolidayDate] = true,
                    [BodyWasExpanded] = true,
                    [PostCategory] = true,
                    [_ResourceType] = true,
                    [Duration] = true,
                    [StartDate] = true,
                    [xd_Signature] = true,
                    [MobileContent] = true,
                    [Preview] = true,
                    [HealthRuleType] = true,
                    [ListType] = true,
                    [IMEPos] = true,
                    [Checkmark] = true,
                    [AppEditor] = true,
                    [DocIcon] = true,
                    [ParentVersionString] = true,
                    [HomeAddressPostalCode] = true,
                    [PersonImage] = true,
                    [UserField1] = true,
                    [PreviouslyAssignedTo] = true,
                    [_DCDateModified] = true,
                    [_Identifier] = true,
                    [GUID] = true,
                    [ProgId] = true,
                    [Language] = true,
                    [UserName] = true,
                    [OffsiteParticipantReason] = true,
                    [WorkAddress] = true,
                    [_ModerationComments] = true,
                    [EventType] = true,
                    [Created] = true,
                    [FolderChildCount] = true,
                    [CorrectBodyToShow] = true,
                    [GbwLocation] = true,
                    [InstanceID] = true,
                    [HomePhone] = true,
                    [WhatsNew] = true,
                    [RelatedLinks] = true,
                    [Birthday] = true,
                    [DiscussionLastUpdated] = true,
                    [DisplayTemplateJSTemplateHidden] = true,
                    [WikiField] = true,
                    [Edit] = true,
                    [XSLStyleWPType] = true,
                    [FSObjType] = true,
                    [_EndDate] = true,
                    [ShortestThreadIndex] = true,
                    [ol_EventAddress] = true,
                    [TelexNumber] = true,
                    [DisplayTemplateJSTemplateType] = true,
                    [HealthRuleCheckEnabled] = true,
                    [RequiredField] = true,
                    [IMAddress] = true,
                    [xd_ProgID] = true,
                    [TotalWork] = true,
                    [FileType] = true,
                    [Nickname] = true,
                    [PrimaryNumber] = true,
                    [ImageCreateDate] = true,
                    [NumberOfVacation] = true,
                    [SystemTask] = true,
                    [IsRootPost] = true,
                    [Late] = true,
                    [UserInfoHidden] = true,
                    [Business2Number] = true,
                    [Created_x0020_By] = true,
                    [FormData] = true,
                    [LinkTitle] = true,
                    [IMEDisplay] = true,
                    [Notes] = true,
                    [_SourceUrl] = true,
                    [FileSizeDisplay] = true,
                    [HealthReportSeverityIcon] = true,
                    [ThumbnailOnForm] = true,
                    [WorkPhone] = true,
                    [TaskGroup] = true,
                    [HealthReportRemedy] = true,
                    [EmailCalendarDateStamp] = true,
                    [MoreLink] = true,
                    [_Source] = true,
                    [MetaInfo] = true,
                    [DateCompleted] = true,
                    [Completed] = true,
                    [ItemChildCount] = true,
                    [SelectFilename] = true,
                    [SurveyTitle] = true,
                    [DayOfWeek] = true,
                    [EncodedAbsUrl] = true,
                    [DLC_Description] = true,
                    [QuotedTextWasExpanded] = true,
                    [IsActive] = true,
                    [MUILanguages] = true,
                    [HomeFaxNumber] = true,
                    [Categories] = true,
                    [ImageSize] = true,
                    [HealthReportServices] = true,
                    [HealthReportServers] = true,
                    [Content] = true,
                    [Predecessors] = true,
                    [PreviewExists] = true,
                    [LessLink] = true,
                    [XSLStyleRequiredFields] = true,
                    [IMEUrl] = true,
                    [SkipFeatureDeployment] = true,
                    [AppPackageErrorMessage] = true,
                    [AppPermissionXML] = true,
                    [AppSubtype] = true,
                    [AppSubtypeID] = true,
                    [AppTitleInfo] = true,
                    [AppVersion] = true,
                    [AssetID] = true,
                    [IsAutoHostedApp] = true,
                    [AadApplicationAppId] = true,
                    [AadApplicationId] = true,
                    [AadServicePrincipalId] = true,
                    [SharePointAppCategory] = true,
                    [IsClientSideSolution] = true,
                    [ClientComponentDeveloperMpnId] = true,
                    [ClientComponentDeveloperName] = true,
                    [ClientComponentDeveloperPrivacyUrl] = true,
                    [ClientComponentDeveloperTermsOfUseUrl] = true,
                    [ClientComponentDeveloperWebsiteUrl] = true,
                    [TenantWideExtensionComponentId] = true,
                    [ClientComponentId] = true,
                    [ClientComponentManifest] = true,
                    [ClientComponentName] = true,
                    [TenantWideExtensionComponentProperties] = true,
                    [ClientComponentType] = true,
                    [ClientComponentConnectedTeamsAppId] = true,
                    [ContainsTeamsManifest] = true,
                    [ContainsTenantWideExtension] = true,
                    [ContainsVivaManifest] = true,
                    [IsClientSideSolutionCurrentVersionDeployed] = true,
                    [IsDefaultAppMetadataLocale] = true,
                    [IsClientSideSolutionDeployed] = true,
                    [AppDescription] = true,
                    [TenantWideExtensionDisabled] = true,
                    [Disabled] = true,
                    [IsAppPackageEnabled] = true,
                    [ExternalContentDomains] = true,
                    [FeatureDescription] = true,
                    [FeatureId] = true,
                    [FeatureName] = true,
                    [FeatureScope] = true,
                    [IsFeaturedApp] = true,
                    [IsClientSideSolutionGeneratedCardDesign] = true,
                    [ClientComponentIsGeneratedCardDesign] = true,
                    [TenantWideExtensionHostProperties] = true,
                    [AppThumbnailURL] = true,
                    [AppImageURL1] = true,
                    [AppImageURL2] = true,
                    [AppImageURL3] = true,
                    [AppImageURL4] = true,
                    [AppImageURL5] = true,
                    [IsolatedDomain] = true,
                    [TenantWideExtensionListTemplate] = true,
                    [TenantWideExtensionLocation] = true,
                    [AppMetadataLocale] = true,
                    [AppCatalog_IsNewVersionAvailable] = true,
                    [PackageDefaultSkipFeatureDeployment] = true,
                    [AppPackageHash] = true,
                    [AppPrerequisitesXML] = true,
                    [PreviouslySetClientComponentSkipFeatureDeployment] = true,
                    [AppProductID] = true,
                    [IsProviderHostedClientSideSolution] = true,
                    [AppPublisher] = true,
                    [ReturnIfCustomScriptDisabled] = true,
                    [TenantWideExtensionSequence] = true,
                    [AppShortDescription] = true,
                    [ClientComponentSkipFeatureDeployment] = true,
                    [SolutionId] = true,
                    [SolutionName] = true,
                    [SolutionVersion] = true,
                    [StoreAssetId] = true,
                    [AppSupportURL] = true,
                    [AppSupportedLocales] = true,
                    [ComponentSupportsFullPage] = true,
                    [ComponentSupportsTeamsTab] = true,
                    [SupportsTeamsTabs] = true,
                    [ComponentSupportsWebPart] = true,
                    [UniqueSolutionId] = true,
                    [IsValidAppPackage] = true,
                    [AppVideoURL] = true,
                    [WebApiPermissionScopesNote] = true,
                    [TenantWideExtensionWebTemplate] = true,
                    [WebApiPermissionRequests] = true
                };
            }
            bool flag = false;
            s_dict.TryGetValue(fid, out flag);
            return flag;
        }
    }
}
