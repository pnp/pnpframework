using PnP.Framework;
using PnP.Framework.Diagnostics;
using PnP.Framework.Enums;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.Async;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that holds the file and folder methods
    /// </summary>
    public static partial class FileFolderExtensions
    {
        /// <summary>
        /// Approves a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to approve</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static void ApproveFile(this Web web, string serverRelativeUrl, string comment)
        {
            Task.Run(() => web.ApproveFileImplementation(serverRelativeUrl, comment)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Approves a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to approve</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static async Task ApproveFileAsync(this Web web, string serverRelativeUrl, string comment)
        {
            await new SynchronizationContextRemover();
            await web.ApproveFileImplementation(serverRelativeUrl, comment);
        }

        /// <summary>
        /// Approves a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to approve</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        private static async Task ApproveFileImplementation(this Web web, string serverRelativeUrl, string comment)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            web.Context.Load(file, x => x.Exists, x => x.CheckOutType);
            await web.Context.ExecuteQueryRetryAsync();

            if (file.Exists)
            {
                file.Approve(comment);
                await web.Context.ExecuteQueryRetryAsync();
            }
        }

        /// <summary>
        /// Checks in a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkin</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static void CheckInFile(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            Task.Run(() => web.CheckInFileImplementation(serverRelativeUrl, checkinType, comment)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks in a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkin</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static async Task CheckInFileAsync(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            await new SynchronizationContextRemover();
            await web.CheckInFileImplementation(serverRelativeUrl, checkinType, comment);
        }

        /// <summary>
        /// Checks in a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkin</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static async Task CheckInFileImplementation(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            var scope = new ConditionalScope(web.Context, () => !file.ServerObjectIsNull.Value && file.Exists && file.CheckOutType != CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file);
            }
            await web.Context.ExecuteQueryAsync();

            if (scope.TestResult.Value)
            {
                file.CheckIn(comment, checkinType);
                await web.Context.ExecuteQueryAsync();
            }
        }

        /// <summary>
        /// Checks out a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkout</param>
        public static void CheckOutFile(this Web web, string serverRelativeUrl)
        {
            Task.Run(() => web.CheckOutFileImplementation(serverRelativeUrl)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks out a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkout</param>
        public static async Task CheckOutFileAsync(this Web web, string serverRelativeUrl)
        {
            await new SynchronizationContextRemover();
            await web.CheckOutFileImplementation(serverRelativeUrl);
        }

        /// <summary>
        /// Checks out a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkout</param>
        private static async Task CheckOutFileImplementation(this Web web, string serverRelativeUrl)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            await web.Context.ExecuteQueryAsync();

            var scope = new ConditionalScope(web.Context, () => !file.ServerObjectIsNull.Value && file.Exists && file.CheckOutType == CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file);
            }
            await web.Context.ExecuteQueryAsync();

            if (scope.TestResult.Value)
            {
                file.CheckOut();
                await web.Context.ExecuteQueryAsync();
            }
        }

        /// <summary>
        /// Discard changes to a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to dicard changes</param>
        public static void UndoCheckOutFile(this Web web, string serverRelativeUrl)
        {
            Task.Run(() => web.UndoCheckOutFileImplementation(serverRelativeUrl)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Discard changes to a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to dicard changes</param>
        public static async Task UndoCheckOutFileAsync(this Web web, string serverRelativeUrl)
        {
            await new SynchronizationContextRemover();
            await web.UndoCheckOutFileImplementation(serverRelativeUrl);
        }

        /// <summary>
        /// Discard changes to a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to dicard changes</param>
        private static async Task UndoCheckOutFileImplementation(this Web web, string serverRelativeUrl)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            await web.Context.ExecuteQueryAsync();

            var scope = new ConditionalScope(web.Context, () => !file.ServerObjectIsNull.Value && file.Exists && file.CheckOutType != CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file);
            }
            await web.Context.ExecuteQueryAsync();

            if (scope.TestResult.Value)
            {
                file.UndoCheckOut();
                web.Context.ExecuteQueryRetry();
            }
        }

        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        /// <summary>
        /// Creates a new document set as a child of an existing folder, with the specified content type ID.
        /// </summary>
        /// <param name="folder">Folder of the document set</param>
        /// <param name="documentSetName">Name of the document set</param>
        /// <param name="contentTypeId">Content type of the document set</param>
        /// <returns>The created Folder representing the document set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <example>
        ///     var setContentType = list.BestMatchContentTypeId(BuiltInContentTypeId.DocumentSet);
        ///     var set1 = list.RootFolder.CreateDocumentSet("Set 1", setContentType);
        /// </example>
        /// </remarks>
        public static Folder CreateDocumentSet(this Folder folder, string documentSetName, ContentTypeId contentTypeId)
        {
            return Task.Run(() => folder.CreateDocumentSetImplementation(documentSetName, contentTypeId)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates a new document set as a child of an existing folder, with the specified content type ID.
        /// </summary>
        /// <param name="folder">Folder of the document set</param>
        /// <param name="documentSetName">Name of the document set</param>
        /// <param name="contentTypeId">Content type of the document set</param>
        /// <returns>The created Folder representing the document set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <example>
        ///     var setContentType = list.BestMatchContentTypeId(BuiltInContentTypeId.DocumentSet);
        ///     var set1 = list.RootFolder.CreateDocumentSet("Set 1", setContentType);
        /// </example>
        /// </remarks>
        public static async Task<Folder> CreateDocumentSetAsync(this Folder folder, string documentSetName, ContentTypeId contentTypeId)
        {
            await new SynchronizationContextRemover();
            return await folder.CreateDocumentSetImplementation(documentSetName, contentTypeId);
        }

        /// <summary>
        /// Creates a new document set as a child of an existing folder, with the specified content type ID.
        /// </summary>
        /// <param name="folder">Folder of the document set</param>
        /// <param name="documentSetName">Name of the document set</param>
        /// <param name="contentTypeId">Content type of the document set</param>
        /// <returns>The created Folder representing the document set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <example>
        ///     var setContentType = list.BestMatchContentTypeId(BuiltInContentTypeId.DocumentSet);
        ///     var set1 = list.RootFolder.CreateDocumentSet("Set 1", setContentType);
        /// </example>
        /// </remarks>
        private static async Task<Folder> CreateDocumentSetImplementation(this Folder folder, string documentSetName, ContentTypeId contentTypeId)
        {
            if (folder == null) { throw new ArgumentNullException(nameof(folder)); }
            if (documentSetName == null) { throw new ArgumentNullException(nameof(documentSetName)); }
            if (contentTypeId == null) { throw new ArgumentNullException(nameof(contentTypeId)); }

            if (documentSetName.ContainsInvalidUrlChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateDocumentSet_The_argument_must_be_a_single_document_set_name_and_cannot_contain_path_characters_, nameof(documentSetName));
            }

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FieldAndContentTypeExtensions_CreateDocumentSet, documentSetName);

            var result = DocumentSet.DocumentSet.Create(folder.Context, folder, documentSetName, contentTypeId);
            await folder.Context.ExecuteQueryAsync();
            var fullUri = new Uri(result.Value);
            var serverRelativeUrl = fullUri.AbsolutePath;

            var ctx = folder.Context as ClientContext;
            var resourcePath = ResourcePath.FromDecodedUrl(serverRelativeUrl);
            var documentSetFolder = ctx.Web.GetFolderByServerRelativePath(resourcePath);

            return documentSetFolder;
        }
        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folderName">Folder name to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder ConvertFolderToDocumentSet(this List list, string folderName)
        {
            var folder = Task.Run(() => list.RootFolder.ResolveSubFolderImplementation(folderName)).GetAwaiter().GetResult();
            if (folder == null) throw new ArgumentException(CoreResources.FileFolderExtensions_FolderMissing);
            return Task.Run(() => list.ConvertFolderToDocumentSetImplementation(folder)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folderName">Folder name to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<Folder> ConvertFolderToDocumentSetAsync(this List list, string folderName)
        {
            await new SynchronizationContextRemover();
            var folder = await list.RootFolder.ResolveSubFolderImplementation(folderName);
            if (folder == null) throw new ArgumentException(CoreResources.FileFolderExtensions_FolderMissing);
            return await list.ConvertFolderToDocumentSetImplementation(folder);
        }

        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folder">Folder to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder ConvertFolderToDocumentSet(this List list, Folder folder)
        {
            return Task.Run(() => list.ConvertFolderToDocumentSetImplementation(folder)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folder">Folder to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<Folder> ConvertFolderToDocumentSetAsync(this List list, Folder folder)
        {
            await new SynchronizationContextRemover();
            return await list.ConvertFolderToDocumentSetImplementation(folder);
        }

        /// <summary>
        /// Internal implementation of the Folder conversion to Document set
        /// </summary>
        /// <param name="list">Library in which the folder exists</param>
        /// <param name="folder">Folder to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        private static async Task<Folder> ConvertFolderToDocumentSetImplementation(this List list, Folder folder)
        {
            list.EnsureProperties(l => l.ContentTypes.Include(c => c.StringId));
            folder.Context.Load(folder.ListItemAllFields, l => l["ContentTypeId"]);
            folder.Context.ExecuteQueryRetry();
            var listItem = folder.ListItemAllFields;

            // If already a document set, just return the folder
            if (listItem["ContentTypeId"].ToString().StartsWith(BuiltInContentTypeId.DocumentSet)) return folder;
            listItem["ContentTypeId"] = BuiltInContentTypeId.DocumentSet;

            // Add missing properties            
            listItem["HTML_x0020_File_x0020_Type"] = "Sharepoint.DocumentSet";
            folder.Properties["docset_LastRefresh"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss");
            folder.Properties["vti_contenttypeorder"] = string.Join(",", list.ContentTypes.ToList().Where(c => c.StringId.StartsWith(BuiltInContentTypeId.Document + "00"))?.Select(c => c.StringId));

            listItem.Update();
            folder.Update();
            await list.Context.ExecuteQueryRetryAsync();
            folder = await list.RootFolder.ResolveSubFolderImplementation(folder.Name);

            //Refresh Folder, otherwise 'Version conflict' error might be thrown on changing properties
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name as a child of the Web. 
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created Folder, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder CreateFolder(this Web web, string folderName)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = Task.Run(() => CreateFolderImplementation(folderCollection, folderName)).GetAwaiter().GetResult();
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name as a child of the Web. 
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created Folder, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<Folder> CreateFolderAsync(this Web web, string folderName)
        {
            await new SynchronizationContextRemover();
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = await CreateFolderImplementation(folderCollection, folderName);
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// <example>
        ///     var folder = list.RootFolder.CreateFolder("new-folder");
        /// </example>
        /// </remarks>
        public static Folder CreateFolder(this Folder parentFolder, string folderName)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = Task.Run(() => CreateFolderImplementation(folderCollection, folderName, parentFolder)).GetAwaiter().GetResult();
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// <example>
        ///     var folder = list.RootFolder.CreateFolder("new-folder");
        /// </example>
        /// </remarks>
        public static async Task<Folder> CreateFolderAsync(this Folder parentFolder, string folderName)
        {
            await new SynchronizationContextRemover();
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = await CreateFolderImplementation(folderCollection, folderName, parentFolder);
            return folder;
        }

        private static async Task<Folder> CreateFolderImplementation(FolderCollection folderCollection, string folderName, Folder parentFolder = null, params Expression<Func<Folder, object>>[] expressions)
        {
            ClientContext context = null;
            if (parentFolder != null)
            {
                context = parentFolder.Context as ClientContext;
            }

            List parentList = null;

            if (parentFolder != null)
            {
                parentFolder.EnsureProperty(p => p.Properties);
                if (parentFolder.Properties.FieldValues.ContainsKey("vti_listname") && context != null)
                {
                    Guid parentListId = Guid.Parse((String)parentFolder.Properties.FieldValues["vti_listname"]);
                    parentList = context.Web.Lists.GetById(parentListId);
                    context.Load(parentList, l => l.BaseType, l => l.Title);
                    await context.ExecuteQueryRetryAsync();
                }
            }

            if (parentList == null)
            {
                // Create folder for library or common URL path
                FolderCollectionAddParameters folderAddParameters = new FolderCollectionAddParameters
                {
                    Overwrite = true
                };

                ResourcePath resourcePath = ResourcePath.FromDecodedUrl(folderName);

                var newFolder = folderCollection.AddUsingPath(resourcePath, folderAddParameters);
                if (expressions != null && expressions.Any())
                {
                    folderCollection.Context.Load(newFolder, expressions);
                }
                else
                {
                    folderCollection.Context.Load(newFolder);
                }
                await folderCollection.Context.ExecuteQueryRetryAsync();
                return newFolder;
            }
            else
            {
                // Create folder for generic list                
                parentFolder.EnsureProperty(f => f.ServerRelativePath);

                ListItemCreationInformationUsingPath newFolderInfo = new ListItemCreationInformationUsingPath
                {
                    LeafName = ResourcePath.FromDecodedUrl(folderName),
                    UnderlyingObjectType = FileSystemObjectType.Folder,
                    FolderPath = ResourcePath.FromDecodedUrl(parentFolder.ServerRelativePath.DecodedUrl)
                };

                ListItem newFolderItem = parentList.AddItemUsingPath(newFolderInfo);
                newFolderItem["Title"] = folderName;
                newFolderItem.Update();
                await context.ExecuteQueryRetryAsync();

                // Get the newly created folder
                var folderPath = parentFolder.ServerRelativePath.DecodedUrl.TrimEnd(new char[] { '/' }) + "/" + folderName;
                var newFolder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(folderPath));
                // Ensure all properties are loaded (to be compatible with the previous implementation)
                if (expressions != null && expressions.Any())
                {
                    context.Load(newFolder, expressions);
                }
                else
                {
                    context.Load(newFolder);
                }
                await context.ExecuteQueryRetryAsync();
                return (newFolder);
            }
        }

        /// <summary>
        /// Checks if a specific folder exists
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeFolderUrl">Folder to check</param>
        /// <returns>Returns true if folder exists</returns>
        public static bool DoesFolderExists(this Web web, string serverRelativeFolderUrl)
        {
            return Task.Run(() => DoesFolderExistImplementation(web, serverRelativeFolderUrl)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks if a specific folder exists
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeFolderUrl">Folder to check</param>
        /// <returns>Returns true if folder exists</returns>
        public static async Task<bool> DoesFolderExistsAsync(this Web web, string serverRelativeFolderUrl)
        {
            await new SynchronizationContextRemover();
            return await DoesFolderExistImplementation(web, serverRelativeFolderUrl);
        }

        private static async Task<bool> DoesFolderExistImplementation(this Web web, string serverRelativeFolderUrl)
        {
            Folder folder = web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeFolderUrl));

            web.Context.Load(folder);
            bool exists = false;

            try
            {
                await web.Context.ExecuteQueryRetryAsync();
                exists = true;
            }
            catch
            {
                return false;
            }

            return exists;
        }

        /// <summary>
        /// Ensure that the folder structure is created. This also ensures hierarchy of folders.
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub site</param>
        /// <param name="parentFolder">Parent folder</param>
        /// <param name="folderPath">Folder path</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The folder structure</returns>
        public static Folder EnsureFolder(this Web web, Folder parentFolder, string folderPath, params Expression<Func<Folder, object>>[] expressions)
        {
            return Task.Run(() => web.EnsureFolderImplementation(parentFolder, folderPath, expressions)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Ensure that the folder structure is created. This also ensures hierarchy of folders.
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub site</param>
        /// <param name="parentFolder">Parent folder</param>
        /// <param name="folderPath">Folder path</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The folder structure</returns>
        public static async Task<Folder> EnsureFolderAsync(this Web web, Folder parentFolder, string folderPath, params Expression<Func<Folder, object>>[] expressions)
        {
            await new SynchronizationContextRemover();
            return await web.EnsureFolderImplementation(parentFolder, folderPath, expressions);
        }

        /// <summary>
        /// Ensure that the folder structure is created. This also ensures hierarchy of folders.
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub site</param>
        /// <param name="parentFolder">Parent folder</param>
        /// <param name="folderPath">Folder path</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The folder structure</returns>
        public static async Task<Folder> EnsureFolderImplementation(this Web web, Folder parentFolder, string folderPath, params Expression<Func<Folder, object>>[] expressions)
        {
            await web.EnsurePropertiesAsync(w => w.ServerRelativeUrl);
            await parentFolder.EnsurePropertiesAsync(f => f.ServerRelativeUrl);

            var parentWebRelativeUrl = parentFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length);
            var webRelativeUrl = parentWebRelativeUrl + (parentWebRelativeUrl.EndsWith("/") ? "" : "/") + folderPath;
            return await web.EnsureFolderPathImplementation(webRelativeUrl, expressions: expressions);
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site, and if it does not exist creates it.
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolder(this Web web, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = Task.Run(() => EnsureFolderImplementation(folderCollection, folderName, expressions: expressions)).GetAwaiter().GetResult();
            return folder;
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site, and if it does not exist creates it.
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<Folder> EnsureFolderAsync(this Web web, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            await new SynchronizationContextRemover();
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = await EnsureFolderImplementation(folderCollection, folderName, expressions: expressions);
            return folder;
        }

        /// <summary>
        /// Checks if the subfolder exists, and if it does not exist creates it.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolder(this Folder parentFolder, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = Task.Run(() => EnsureFolderImplementation(folderCollection, folderName, parentFolder, expressions)).GetAwaiter().GetResult();
            return folder;
        }

        /// <summary>
        /// Checks if the subfolder exists, and if it does not exist creates it.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<Folder> EnsureFolderAsync(this Folder parentFolder, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            await new SynchronizationContextRemover();
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = await EnsureFolderImplementation(folderCollection, folderName, parentFolder, expressions);
            return folder;
        }

        private static async Task<Folder> EnsureFolderImplementation(FolderCollection folderCollection, string folderName, Folder parentFolder = null, params Expression<Func<Folder, object>>[] expressions)
        {
            Folder folder = null;
            if (expressions != null && expressions.Any())
            {
                folderCollection.Context.Load(folderCollection, fc => fc.IncludeWithDefaultProperties(expressions));
            }
            else
            {
                folderCollection.Context.Load(folderCollection);
            }
            await folderCollection.Context.ExecuteQueryRetryAsync();

            foreach (Folder existingFolder in folderCollection)
            {
                if (string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    folder = existingFolder;
                    break;
                }
            }

            if (folder == null)
            {
                folder = await CreateFolderImplementation(folderCollection, folderName, parentFolder, expressions);
            }

            return folder;
        }
        /// <summary>
        /// Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
        /// </summary>
        /// <param name="web">Web to check for the specified folder</param>
        /// <param name="webRelativeUrl">Path to the folder, relative to the web site</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// If the specified path is inside an existing list, then the folder is created inside that list.
        /// </para>
        /// <para>
        /// Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolderPath(this Web web, string webRelativeUrl, params Expression<Func<Folder, object>>[] expressions)
        {
            return Task.Run(() => web.EnsureFolderPathImplementation(webRelativeUrl, expressions)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
        /// </summary>
        /// <param name="web">Web to check for the specified folder</param>
        /// <param name="webRelativeUrl">Path to the folder, relative to the web site</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// If the specified path is inside an existing list, then the folder is created inside that list.
        /// </para>
        /// <para>
        /// Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
        /// </para>
        /// </remarks>
        public static async Task<Folder> EnsureFolderPathAsync(this Web web, string webRelativeUrl, params Expression<Func<Folder, object>>[] expressions)
        {
            await new SynchronizationContextRemover();
            return await web.EnsureFolderPathImplementation(webRelativeUrl, expressions);
        }

        /// <summary>
        /// Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
        /// </summary>
        /// <param name="web">Web to check for the specified folder</param>
        /// <param name="webRelativeUrl">Path to the folder, relative to the web site</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// If the specified path is inside an existing list, then the folder is created inside that list.
        /// </para>
        /// <para>
        /// Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
        /// </para>
        /// </remarks>
        private static async Task<Folder> EnsureFolderPathImplementation(this Web web, string webRelativeUrl, params Expression<Func<Folder, object>>[] expressions)
        {

            if (webRelativeUrl == null) { throw new ArgumentNullException(nameof(webRelativeUrl)); }

            if (webRelativeUrl.EndsWith("."))
            {
                throw new Exception("Folder names cannot end on a period (.).");
            }
            //Web root folder should be returned if webRelativeUrl is empty
            if (webRelativeUrl.Length != 0 && string.IsNullOrWhiteSpace(webRelativeUrl)) { throw new ArgumentException(CoreResources.FileFolderExtensions_EnsureFolderPath_Folder_URL_is_required_, nameof(webRelativeUrl)); }

            await web.EnsurePropertyAsync(w => w.ServerRelativeUrl);

            var folderServerRelativeUrl = UrlUtility.Combine(web.ServerRelativeUrl, webRelativeUrl, "/");

            // Check if folder is inside a list. We need to exclude Solution Catalog since users that are not site owners will not have access to the root folder and we cannot anyhow create folders there.
            var listCollection = web.Context.LoadQuery(web.Lists.Where(l => l.BaseTemplate != (int)ListTemplateType.SolutionCatalog).Include(l => l.RootFolder));

            await web.Context.ExecuteQueryRetryAsync();

            List containingList = null;

            foreach (var list in listCollection)
            {
                if (folderServerRelativeUrl.StartsWith(UrlUtility.Combine(list.RootFolder.ServerRelativeUrl, "/"), StringComparison.InvariantCultureIgnoreCase))
                {
                    // Load fields from the list
                    containingList = list;
                    break;
                }
            }

            // Start either at the root of the list or web
            string locationType = null;
            string listUrl = string.Empty;
            IEnumerable<Field> titleField = null;
            string rootUrl = null;

            Folder currentFolder = null;
            if (containingList == null)
            {
                locationType = "Web";
                currentFolder = await web.EnsurePropertyAsync(w => w.RootFolder);
            }
            else
            {
                locationType = "List";
                currentFolder = containingList.RootFolder;
                listUrl = containingList.RootFolder.ServerRelativeUrl;

                titleField = web.Context.LoadQuery(containingList.Fields.Where(f => f.Id == BuiltInFieldId.Title));
                await web.Context.ExecuteQueryRetryAsync();
            }
            rootUrl = currentFolder.ServerRelativeUrl;

            // Get remaining parts of the path and split
            var folderRootRelativeUrl = folderServerRelativeUrl.Substring(currentFolder.ServerRelativeUrl.Length);
            var childFolderNames = folderRootRelativeUrl.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentCount = 0;

            foreach (var folderName in childFolderNames)
            {
                currentCount++;

                // Find next part of the path
                var folderCollection = currentFolder.Folders;
                folderCollection.Context.Load(folderCollection);
                await folderCollection.Context.ExecuteQueryRetryAsync();

                Folder nextFolder = null;
                foreach (Folder existingFolder in folderCollection)
                {
                    //WebUtility.UrlDecode removes + from folderName which leads to invalid compare if folderName was not UrlEncoded --> replaced with Uri.UnescapeDatastring
                    if (string.Equals(existingFolder.Name, Uri.UnescapeDataString(folderName), StringComparison.InvariantCultureIgnoreCase) || string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        nextFolder = existingFolder;
                        break;
                    }
                }

                // Or create it
                if (nextFolder == null)
                {
                    var createPath = string.Join("/", childFolderNames, 0, currentCount);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_CreateFolder0Under12, createPath, locationType, rootUrl);
                    if (locationType == "List")
                    {
                        createPath = createPath.Substring(0, createPath.Length - folderName.Length).TrimEnd('/');

                        var newFolderInfo = new ListItemCreationInformationUsingPath
                        {
                            UnderlyingObjectType = FileSystemObjectType.Folder,
                            FolderPath = ResourcePath.FromDecodedUrl(UrlUtility.Combine(listUrl, createPath)),
                            LeafName = ResourcePath.FromDecodedUrl(folderName)
                        };

                        ListItem newFolderItem = containingList.AddItemUsingPath(newFolderInfo);

                        if (titleField != null && titleField.Any())
                        {
                            newFolderItem["Title"] = folderName;
                            newFolderItem.Update();
                        }
                        
                        containingList.Context.Load(newFolderItem);
                        await containingList.Context.ExecuteQueryRetryAsync();

                        nextFolder = web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(UrlUtility.Combine(listUrl, createPath, folderName)));
                        containingList.Context.Load(nextFolder);
                        await containingList.Context.ExecuteQueryRetryAsync();
                    }
                    else
                    {
                        nextFolder = folderCollection.Add(folderName);
                        folderCollection.Context.Load(nextFolder);
                        await folderCollection.Context.ExecuteQueryRetryAsync();
                    }
                }

                currentFolder = nextFolder;
            }
            if (expressions != null && expressions.Any())
            {
                web.Context.Load(currentFolder, expressions);
                await web.Context.ExecuteQueryRetryAsync();
            }
            return currentFolder;
        }

        /// <summary>
        /// Finds files in the web. Can be slow.
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this Web web, string match)
        {
            Folder rootFolder = web.RootFolder;
            match = WildcardToRegex(match);
            return Task.Run(() => ParseFiles(rootFolder, match, web.Context as ClientContext)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Finds files in the web. Can be slow.
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static async Task<List<File>> FindFilesAsync(this Web web, string match)
        {
            await new SynchronizationContextRemover();
            Folder rootFolder = web.RootFolder;
            match = WildcardToRegex(match);
            return await ParseFiles(rootFolder, match, web.Context as ClientContext);
        }

        /// <summary>
        /// Find files in the list, Can be slow.
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this List list, string match)
        {
            Folder rootFolder = list.EnsureProperty(l => l.RootFolder);

            match = WildcardToRegex(match);
            return Task.Run(() => ParseFiles(rootFolder, match, list.Context as ClientContext)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Find files in the list, Can be slow.
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static async Task<List<File>> FindFilesAsync(this List list, string match)
        {
            await new SynchronizationContextRemover();
            Folder rootFolder = list.EnsureProperty(l => l.RootFolder);

            match = WildcardToRegex(match);
            return await ParseFiles(rootFolder, match, list.Context as ClientContext);
        }

        /// <summary>
        /// Find files in a specific folder
        /// </summary>
        /// <param name="folder">The folder to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this Folder folder, string match)
        {
            match = WildcardToRegex(match);
            return Task.Run(() => ParseFiles(folder, match, folder.Context as ClientContext)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Find files in a specific folder
        /// </summary>
        /// <param name="folder">The folder to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static async Task<List<File>> FindFilesAsync(this Folder folder, string match)
        {
            await new SynchronizationContextRemover();
            match = WildcardToRegex(match);
            return await ParseFiles(folder, match, folder.Context as ClientContext);
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static bool FolderExists(this Web web, string folderName)
        {
            var folderCollection = web.Folders;
            var exists = Task.Run(() => FolderExistsImplementation(folderCollection, folderName)).GetAwaiter().GetResult();
            return exists;
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<bool> FolderExistsAsync(this Web web, string folderName)
        {
            await new SynchronizationContextRemover();
            var folderCollection = web.Folders;
            var exists = await FolderExistsImplementation(folderCollection, folderName);
            return exists;
        }

        /// <summary>
        /// Checks if the subfolder exists.
        /// </summary>
        /// <param name="parentFolder">Parent folder to check for the named subfolder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static bool FolderExists(this Folder parentFolder, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var exists = Task.Run(() => FolderExistsImplementation(folderCollection, folderName)).GetAwaiter().GetResult();
            return exists;
        }

        /// <summary>
        /// Checks if the subfolder exists.
        /// </summary>
        /// <param name="parentFolder">Parent folder to check for the named subfolder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static async Task<bool> FolderExistsAsync(this Folder parentFolder, string folderName)
        {
            await new SynchronizationContextRemover();
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var exists = await FolderExistsImplementation(folderCollection, folderName);
            return exists;
        }

        private static async Task<bool> FolderExistsImplementation(FolderCollection folderCollection, string folderName)
        {
            if (folderCollection == null)
            {
                throw new ArgumentNullException(nameof(folderCollection));
            }

            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            folderCollection.Context.Load(folderCollection);
            await folderCollection.Context.ExecuteQueryRetryAsync();

            foreach (Folder folder in folderCollection)
            {
                if (folder.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
        /// <summary>
        /// Returns a file as string
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <returns>The file contents as a string</returns>
        public static string GetFileAsString(this Web web, string serverRelativeUrl)
        {
            return Task.Run(() => web.GetFileAsStringImplementation(serverRelativeUrl)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns a file as string
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <returns>The file contents as a string</returns>
        public static async Task<string> GetFileAsStringAsync(this Web web, string serverRelativeUrl)
        {
            await new SynchronizationContextRemover();
            return await web.GetFileAsStringImplementation(serverRelativeUrl);
        }

        /// <summary>
        /// Returns a file as string
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <returns>The file contents as a string</returns>
        private static async Task<string> GetFileAsStringImplementation(this Web web, string serverRelativeUrl)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            web.Context.Load(file);
            await web.Context.ExecuteQueryRetryAsync();
            ClientResult<Stream> stream = file.OpenBinaryStream();
            await web.Context.ExecuteQueryRetryAsync();

            string returnString = string.Empty;
            using (Stream memStream = new MemoryStream())
            {
                CopyStream(stream.Value, memStream);
                memStream.Position = 0;
                using (var reader = new StreamReader(memStream))
                {
                    returnString = reader.ReadToEnd();
                }
            }

            return returnString;
        }

        private static async Task<List<File>> ParseFiles(Folder folder, string match, ClientContext context)
        {
            var foundFiles = new List<File>();
            FileCollection files = folder.Files;
            context.Load(files, fs => fs.Include(f => f.ServerRelativeUrl, f => f.Name, f => f.Title, f => f.TimeCreated, f => f.TimeLastModified));
            context.Load(folder.Folders);
            await context.ExecuteQueryRetryAsync();

            foreach (File file in files)
            {
                if (Regex.IsMatch(file.Name, match, RegexOptions.IgnoreCase))
                {
                    foundFiles.Add(file);
                }
            }

            foreach (Folder subfolder in folder.Folders)
            {
                foundFiles.AddRange(await ParseFiles(subfolder, match, context));
            }
            return foundFiles;
        }
        /// <summary>
        /// Publishes a file existing on a server URL
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">the server relative URL of the file to publish</param>
        /// <param name="comment">Comment recorded with the publish action</param>
        public static void PublishFile(this Web web, string serverRelativeUrl, string comment)
        {
            Task.Run(() => web.PublishFileImplementation(serverRelativeUrl, comment)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Publishes a file existing on a server URL
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">the server relative URL of the file to publish</param>
        /// <param name="comment">Comment recorded with the publish action</param>
        public static async Task PublishFileAsync(this Web web, string serverRelativeUrl, string comment)
        {
            await new SynchronizationContextRemover();
            await web.PublishFileImplementation(serverRelativeUrl, comment);
        }

        /// <summary>
        /// Publishes a file existing on a server URL
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">the server relative URL of the file to publish</param>
        /// <param name="comment">Comment recorded with the publish action</param>
        private static async Task PublishFileImplementation(this Web web, string serverRelativeUrl, string comment)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));

            web.Context.Load(file, x => x.Exists, x => x.CheckOutType);
            await web.Context.ExecuteQueryRetryAsync();

            if (file.Exists)
            {
                file.Publish(comment);
                await web.Context.ExecuteQueryRetryAsync();
            }
        }
        /// <summary>
        /// Gets a folder with a given name in a given <see cref="Microsoft.SharePoint.Client.Folder"/>
        /// </summary>
        /// <param name="folder"><see cref="Microsoft.SharePoint.Client.Folder"/> in which to search for</param>
        /// <param name="folderName">Name of the folder to search for</param>
        /// <returns>The found <see cref="Microsoft.SharePoint.Client.Folder"/> if available, null otherwise</returns>
        public static Folder ResolveSubFolder(this Folder folder, string folderName)
        {
            return Task.Run(() => folder.ResolveSubFolderImplementation(folderName)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets a folder with a given name in a given <see cref="Microsoft.SharePoint.Client.Folder"/>
        /// </summary>
        /// <param name="folder"><see cref="Microsoft.SharePoint.Client.Folder"/> in which to search for</param>
        /// <param name="folderName">Name of the folder to search for</param>
        /// <returns>The found <see cref="Microsoft.SharePoint.Client.Folder"/> if available, null otherwise</returns>
        public static async Task<Folder> ResolveSubFolderAsync(this Folder folder, string folderName)
        {
            await new SynchronizationContextRemover();
            return await folder.ResolveSubFolderImplementation(folderName);
        }

        /// <summary>
        /// Gets a folder with a given name in a given <see cref="Microsoft.SharePoint.Client.Folder"/>
        /// </summary>
        /// <param name="folder"><see cref="Microsoft.SharePoint.Client.Folder"/> in which to search for</param>
        /// <param name="folderName">Name of the folder to search for</param>
        /// <returns>The found <see cref="Microsoft.SharePoint.Client.Folder"/> if available, null otherwise</returns>
        private static async Task<Folder> ResolveSubFolderImplementation(this Folder folder, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            folder.Context.Load(folder);
            folder.Context.Load(folder.Folders);
            await folder.Context.ExecuteQueryRetryAsync();

            foreach (Folder subFolder in folder.Folders)
            {
                if (subFolder.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return subFolder;
                }
            }

            return null;
        }

        /// <summary>
        /// Saves a remote file to a local folder
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <param name="localPath">The local folder</param>
        /// <param name="localFileName">The local filename. If null the filename of the file on the server will be used</param>
        /// <param name="fileExistsCallBack">Optional callback function allowing to provide feedback if the file should be overwritten if it exists. The function requests a bool as return value and the string input contains the name of the file that exists.</param>
        public static void SaveFileToLocal(this Web web, string serverRelativeUrl, string localPath, string localFileName = null, Func<string, bool> fileExistsCallBack = null)
        {
            Task.Run(() => web.SaveFileToLocalImplementation(serverRelativeUrl, localPath, localFileName, fileExistsCallBack)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Saves a remote file to a local folder
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <param name="localPath">The local folder</param>
        /// <param name="localFileName">The local filename. If null the filename of the file on the server will be used</param>
        /// <param name="fileExistsCallBack">Optional callback function allowing to provide feedback if the file should be overwritten if it exists. The function requests a bool as return value and the string input contains the name of the file that exists.</param>
        public static async Task SaveFileToLocalAsync(this Web web, string serverRelativeUrl, string localPath, string localFileName = null, Func<string, bool> fileExistsCallBack = null)
        {
            await new SynchronizationContextRemover();
            await web.SaveFileToLocalImplementation(serverRelativeUrl, localPath, localFileName, fileExistsCallBack);
        }

        /// <summary>
        /// Saves a remote file to a local folder
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <param name="localPath">The local folder</param>
        /// <param name="localFileName">The local filename. If null the filename of the file on the server will be used</param>
        /// <param name="fileExistsCallBack">Optional callback function allowing to provide feedback if the file should be overwritten if it exists. The function requests a bool as return value and the string input contains the name of the file that exists.</param>
        public static async Task SaveFileToLocalImplementation(this Web web, string serverRelativeUrl, string localPath, string localFileName = null, Func<string, bool> fileExistsCallBack = null)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            var clientContext = web.Context as ClientContext;
            clientContext.Load(file);
            await clientContext.ExecuteQueryRetryAsync();

            ClientResult<Stream> stream = file.OpenBinaryStream();
            await clientContext.ExecuteQueryRetryAsync();

            var fileOut = Path.Combine(localPath, !string.IsNullOrEmpty(localFileName) ? localFileName : file.Name);

            if (!System.IO.File.Exists(fileOut) || (fileExistsCallBack != null && fileExistsCallBack(fileOut)))
            {
                using (Stream fileStream = new FileStream(fileOut, FileMode.Create))
                {
                    CopyStream(stream.Value, fileStream);
                }
            }
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Name of the file</param>
        /// <param name="localFilePath">Location of the file to be uploaded.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFile(this Folder folder, string fileName, string localFilePath, bool overwriteIfExists)
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));
            if (localFilePath == null)
                throw new ArgumentNullException(nameof(localFilePath));
            if (!System.IO.File.Exists(localFilePath))
                throw new FileNotFoundException("Local file was not found.", localFilePath);

            using (var stream = System.IO.File.OpenRead(localFilePath))
                return Task.Run(() => folder.UploadFileImplementation(fileName, stream, overwriteIfExists)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Name of the file</param>
        /// <param name="localFilePath">Location of the file to be uploaded.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static async Task<File> UploadFileAsync(this Folder folder, string fileName, string localFilePath, bool overwriteIfExists)
        {
            await new SynchronizationContextRemover();
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));
            if (localFilePath == null)
                throw new ArgumentNullException(nameof(localFilePath));
            if (!System.IO.File.Exists(localFilePath))
                throw new FileNotFoundException("Local file was not found.", localFilePath);

            using (var stream = System.IO.File.OpenRead(localFilePath))
                return await folder.UploadFileImplementation(fileName, stream, overwriteIfExists);
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream">A stream object that represents the file.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFile(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            return Task.Run(() => folder.UploadFileImplementation(fileName, stream, overwriteIfExists)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream">A stream object that represents the file.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static async Task<File> UploadFileAsync(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            await new SynchronizationContextRemover();
            return await folder.UploadFileImplementation(fileName, stream, overwriteIfExists);
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream">A stream object that represents the file.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static async Task<File> UploadFileImplementation(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            if (fileName == null)
                throw new ArgumentNullException(nameof(fileName));

            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException(CoreResources.FileFolderExtensions_UploadFile_Destination_file_name_is_required_, nameof(fileName));

            if (fileName.ContainsInvalidFileFolderChars())
                throw new ArgumentException(string.Format(CoreResources.FileFolderExtensions_UploadFile_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_, fileName), nameof(fileName));

            // 10 MB
            int blockSize = 10 * 1024 * 1024;

            File uploadFile = null;
            folder.EnsureProperty(f => f.ServerRelativePath);
            if (stream.Length <= blockSize)
            {
                FileCollectionAddParameters fileCollectionAddParameters = new FileCollectionAddParameters
                {
                    Overwrite = overwriteIfExists
                };
                ResourcePath decodedfileName = ResourcePath.FromDecodedUrl(fileName);
                uploadFile = folder.Files.AddUsingPath(decodedfileName, fileCollectionAddParameters, stream);
                folder.Context.Load(uploadFile);
                await folder.Context.ExecuteQueryRetryAsync();
                // Return the file object for the uploaded file.
                return uploadFile;
            }
            else
            {
                try
                {
                    // Each sliced upload requires a unique ID.
                    Guid uploadId = Guid.NewGuid();
                    ResourcePath decodedfileName = ResourcePath.FromDecodedUrl(fileName);

                    byte[] buffer = new byte[blockSize];
                    byte[] lastBuffer = null;
                    long fileoffset = 0;
                    long totalBytesRead = 0;
                    int bytesRead;
                    bool first = true;
                    bool last = false;

                    // Use large file upload approach.
                    ClientResult<long> bytesUploaded = null;
                    // Read data from file system in blocks.
                    while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        totalBytesRead += bytesRead;

                        // You've reached the end of the file.
                        if (totalBytesRead == stream.Length)
                        {
                            last = true;
                            // Copy to a new buffer that has the correct size.
                            lastBuffer = new byte[bytesRead];
                            Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                        }

                        if (first)
                        {
                            using (MemoryStream contentStream = new MemoryStream())
                            {
                                // Add an empty file.
                                FileCollectionAddParameters fileCollectionAddParameters = new FileCollectionAddParameters
                                {
                                    Overwrite = overwriteIfExists
                                };

                                uploadFile = folder.Files.AddUsingPath(decodedfileName, fileCollectionAddParameters, contentStream);

                                // Start upload by uploading the first slice.
                                using (MemoryStream s = new MemoryStream(buffer))
                                {
                                    Log.Debug(Constants.LOGGING_SOURCE, "Creating file with name '{0}'", decodedfileName);
                                    // Call the start upload method on the first slice.
                                    bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                    await folder.Context.ExecuteQueryRetryAsync();
                                    // fileoffset is the pointer where the next slice will be added.
                                    fileoffset = bytesUploaded.Value;
                                }

                                // You can only start the upload once.
                                first = false;
                            }
                        }
                        else
                        {
                            ResourcePath fileServerRelativePath = ResourcePath.FromDecodedUrl(folder.ServerRelativePath.DecodedUrl + Path.AltDirectorySeparatorChar + decodedfileName.DecodedUrl);
                            uploadFile = (folder.Context as ClientContext).Web.GetFileByServerRelativePath(fileServerRelativePath);
                            if (last)
                            {
                                // Is this the last slice of data?
                                using (MemoryStream s = new MemoryStream(lastBuffer))
                                {
                                    // End sliced upload by calling FinishUpload.
                                    uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                    await folder.Context.ExecuteQueryRetryAsync();
                                }
                            }
                            else
                            {
                                using (MemoryStream s = new MemoryStream(buffer))
                                {
                                    // Continue sliced upload.
                                    bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                    await folder.Context.ExecuteQueryRetryAsync();
                                    // Update fileoffset for the next slice.
                                    fileoffset = bytesUploaded.Value;
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    Log.Error(Constants.LOGGING_SOURCE, "Error creating file with name '{0}'", fileName);
                    throw;
                }
            }

            Log.Debug(Constants.LOGGING_SOURCE, "Created file with name '{0}'", fileName);


            var fileUrl = folder.ServerRelativePath.DecodedUrl + Path.AltDirectorySeparatorChar + fileName;
            var finishedUploadedFile = (folder.Context as ClientContext).Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileUrl));
            folder.Context.Load(finishedUploadedFile);
            await folder.Context.ExecuteQueryRetryAsync();
            // Return the file object for the uploaded file.
            return finishedUploadedFile;

        }

        /// <summary>
        /// Gets a file in a document library.
        /// </summary>
        /// <param name="folder">Folder containing the target file.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>The target file if found, null if no file is found.</returns>
        public static File GetFile(this Folder folder, string fileName)
        {
            return Task.Run(() => folder.GetFileImplementation(fileName)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Gets a file in a document library.
        /// </summary>
        /// <param name="folder">Folder containing the target file.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>The target file if found, null if no file is found.</returns>
        public static async Task<File> GetFileAsync(this Folder folder, string fileName)
        {
            return await folder.GetFileImplementation(fileName);
        }

        /// <summary>
        /// Gets a file in a document library.
        /// </summary>
        /// <param name="folder">Folder containing the target file.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>The target file if found, null if no file is found.</returns>
        private static async Task<File> GetFileImplementation(this Folder folder, string fileName)
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));

            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException(nameof(fileName));

            try
            {
                folder.EnsureProperties(f => f.ServerRelativeUrl);

                var fileServerRelativeUrl = UrlUtility.Combine(folder.ServerRelativeUrl, fileName);
                var context = folder.Context as ClientContext;

                var web = context.Web;

                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
                web.Context.Load(file);
                await web.Context.ExecuteQueryRetryAsync();

                return file;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return null;
                }
                throw;
            }
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localFile">File to validate against.</param>
        public static bool VerifyIfUploadRequired(this File serverFile, string localFile)
        {
            if (localFile == null)
                throw new ArgumentNullException(nameof(localFile));
            if (!System.IO.File.Exists(localFile))
                throw new FileNotFoundException("Local file was not found.", localFile);

            using (var file = System.IO.File.OpenRead(localFile))
                return Task.Run(() => serverFile.VerifyIfUploadRequiredImplementation(file)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localFile">File to validate against.</param>
        public static async Task<bool> VerifyIfUploadRequiredAsync(this File serverFile, string localFile)
        {
            await new SynchronizationContextRemover();
            if (localFile == null)
                throw new ArgumentNullException(nameof(localFile));
            if (!System.IO.File.Exists(localFile))
                throw new FileNotFoundException("Local file was not found.", localFile);

            using (var file = System.IO.File.OpenRead(localFile))
                return await serverFile.VerifyIfUploadRequiredImplementation(file);
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localStream">Stream to validate against.</param>
        /// <returns></returns>
        public static bool VerifyIfUploadRequired(this File serverFile, Stream localStream)
        {
            return Task.Run(() => serverFile.VerifyIfUploadRequiredImplementation(localStream)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localStream">Stream to validate against.</param>
        /// <returns></returns>
        public static async Task<bool> VerifyIfUploadRequiredAsync(this File serverFile, Stream localStream)
        {
            await new SynchronizationContextRemover();
            return await serverFile.VerifyIfUploadRequiredImplementation(localStream);
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localStream">Stream to validate against.</param>
        /// <returns></returns>
        public static async Task<bool> VerifyIfUploadRequiredImplementation(this File serverFile, Stream localStream)
        {
            if (serverFile == null)
                throw new ArgumentNullException(nameof(serverFile));
            if (localStream == null)
                throw new ArgumentNullException(nameof(localStream));

            byte[] serverHash = null;
            var streamResult = serverFile.OpenBinaryStream();
            await serverFile.Context.ExecuteQueryRetryAsync();

            byte[] localHash = null;
            // Hash contents
            using (HashAlgorithm ha = HashAlgorithm.Create("SHA"))
            {

                using (var serverStream = streamResult.Value)
                    serverHash = ha.ComputeHash(serverStream);

                // Check hash (& rewind)
                localHash = ha.ComputeHash(localStream);
                localStream.Position = 0;
            }

            // Compare hash
            var contentsMatch = true;
            for (var index = 0; index < serverHash.Length; index++)
                if (serverHash[index] != localHash[index])
                {
                    contentsMatch = false;
                    break;
                }

            localStream.Position = 0;
            return !contentsMatch;
        }
        /// <summary>
        /// Sets file properties using a dictionary.
        /// </summary>
        /// <param name="file">Target file object.</param>
        /// <param name="properties">Dictionary of properties to set.</param>
        /// <param name="checkoutIfRequired">Check out the file if necessary to set properties.</param>
        public static void SetFileProperties(this File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            Task.Run(() => file.SetFilePropertiesImplementation(properties, checkoutIfRequired)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Sets file properties using a dictionary.
        /// </summary>
        /// <param name="file">Target file object.</param>
        /// <param name="properties">Dictionary of properties to set.</param>
        /// <param name="checkoutIfRequired">Check out the file if necessary to set properties.</param>
        public static async Task SetFilePropertiesAsync(this File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            await new SynchronizationContextRemover();
            await file.SetFilePropertiesImplementation(properties, checkoutIfRequired);
        }

        /// <summary>
        /// Sets file properties using a dictionary.
        /// </summary>
        /// <param name="file">Target file object.</param>
        /// <param name="properties">Dictionary of properties to set.</param>
        /// <param name="checkoutIfRequired">Check out the file if necessary to set properties.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static async Task SetFilePropertiesImplementation(this File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file));
            if (properties == null)
                throw new ArgumentNullException(nameof(properties));

            var changedProperties = new Dictionary<string, string>();
            var changedPropertiesString = new StringBuilder();
            var context = file.Context;

            if (properties != null && properties.Count > 0)
            {
                // Get a reference to the target list, if any
                // and load file item properties
                var parentList = file.ListItemAllFields.ParentList;
                context.Load(parentList, l => l.ForceCheckout);
                context.Load(file.ListItemAllFields);
                context.Load(file.ListItemAllFields.FieldValuesAsText);
                try
                {
                    await context.ExecuteQueryRetryAsync();
                }
                catch (ServerException ex)
                {
                    // If this throws ServerException (does not belong to list), then shouldn't be trying to set properties)
                    // Handling the exception stating the "The object specified does not belong to a list."
                    if (ex.ServerErrorCode != -2113929210)
                    {
                        throw;
                    }
                }

                // Loop through and detect changes first, then, check out if required and apply
                foreach (var kvp in properties)
                {
                    var propertyName = kvp.Key;
                    var propertyValue = kvp.Value;

                    var fieldValues = file.ListItemAllFields.FieldValues;
                    var currentValue = string.Empty;
                    if (file.ListItemAllFields.FieldValues.ContainsKey(propertyName))
                    {
                        currentValue = file.ListItemAllFields.FieldValuesAsText[propertyName];
                    }

                    //LoggingUtility.Internal.TraceVerbose("*** Comparing property '{0}' to current '{1}', new '{2}'", propertyName, currentValue, propertyValue);
                    switch (propertyName.ToUpperInvariant())
                    {
                        case "CONTENTTYPE":
                            {
                                if (!currentValue.Equals(propertyValue, StringComparison.InvariantCultureIgnoreCase) && parentList != null)
                                {
                                    ContentType targetCT = parentList.GetContentTypeByName(propertyValue);
                                    await context.ExecuteQueryRetryAsync();

                                    if (targetCT != null)
                                    {
                                        changedProperties["ContentTypeId"] = targetCT.StringId;
                                        changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                    }
                                    else
                                    {
                                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_SetFileProperties_Error, propertyValue);
                                    }
                                }
                                break;
                            }
                        case "CONTENTTYPEID":
                            {
                                if (!currentValue.Equals(propertyValue, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                break;
                            }
                        case "PUBLISHINGASSOCIATEDCONTENTTYPE":
                            {
                                var testValue = ";#" + currentValue.Replace(", ", ";#") + ";#";
                                if (testValue != propertyValue)
                                {
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                break;
                            }
                        default:
                            {
                                if (currentValue != propertyValue)
                                {
                                    //Console.WriteLine("Setting property '{0}' to '{1}'", propertyName, propertyValue);
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                break;
                            }
                    }
                }

                if (changedProperties.Count > 0)
                {
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_UpdateFile0Properties1, file.Name, changedPropertiesString);
                    var checkOutRequired = false;

                    if (parentList != null)
                    {
                        checkOutRequired = parentList.ForceCheckout;
                    }

                    if (checkoutIfRequired && checkOutRequired && file.CheckOutType == CheckOutType.None)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Checking out file '{0}'", file.Name);
                        file.CheckOut();
                        await context.ExecuteQueryRetryAsync();
                    }

                    Log.Debug(Constants.LOGGING_SOURCE, "Set properties: {0}", file.Name);
                    foreach (var kvp in changedProperties)
                    {
                        var propertyName = kvp.Key;
                        var propertyValue = kvp.Value;

                        Log.Debug(Constants.LOGGING_SOURCE, " {0}={1}", propertyName, propertyValue);
                        file.ListItemAllFields[propertyName] = propertyValue;
                    }
                    file.ListItemAllFields.Update();
                    await context.ExecuteQueryRetryAsync();
                }
            }
        }
        /// <summary>
        /// Publishes a file based on the type of versioning required on the parent library.
        /// </summary>
        /// <param name="file">Target file to publish.</param>
        /// <param name="level">Target publish direction (Draft and Published only apply, Checkout is ignored).</param>
        public static void PublishFileToLevel(this File file, FileLevel level)
        {
            Task.Run(() => file.PublishFileToLevelImplementation(level)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Publishes a file based on the type of versioning required on the parent library.
        /// </summary>
        /// <param name="file">Target file to publish.</param>
        /// <param name="level">Target publish direction (Draft and Published only apply, Checkout is ignored).</param>
        public static async Task PublishFileToLevelAsync(this File file, FileLevel level)
        {
            await new SynchronizationContextRemover();
            await file.PublishFileToLevelImplementation(level);
        }

        /// <summary>
        /// Publishes a file based on the type of versioning required on the parent library.
        /// </summary>
        /// <param name="file">Target file to publish.</param>
        /// <param name="level">Target publish direction (Draft and Published only apply, Checkout is ignored).</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static async Task PublishFileToLevelImplementation(this File file, FileLevel level)
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file));

            var publishingRequired = false;
            var approvalRequired = false;

            if (level == FileLevel.Draft || level == FileLevel.Published)
            {
                var context = file.Context;

                bool normalFile = true;
                // Ensure that ListItemAllFields.ServerObjectIsNull is loaded
                try
                {
                    await file.EnsurePropertiesImplementation<File>(f => f.ListItemAllFields, f => f.CheckOutType, f => f.Name);
                }
                catch
                {
                    // Catch all errors...there's a valid scenario for this failing when this is not a file associated to a listitem
                    normalFile = false;
                }

                // Only access ListItemAllFields if the above load succeeded. If it didn't, accessing it will throw it back in the context, and the next
                // ExecuteQueryRetry will throw a 'The object specified does not belong to a list.' error.
                normalFile = normalFile && (!file.ListItemAllFields.ServerObjectIsNull ?? false); //normal files have listItemAllFields;
                var checkOutRequired = false;
                if (normalFile)
                {
                    var parentList = file.ListItemAllFields.ParentList;
                    context.Load(parentList,
                                l => l.EnableMinorVersions,
                                l => l.EnableModeration,
                                l => l.ForceCheckout);

                    try
                    {
                        context.ExecuteQueryRetry();
                        checkOutRequired = parentList.ForceCheckout;
                        publishingRequired = parentList.EnableMinorVersions; // minor versions implies that the file must be published
                        approvalRequired = parentList.EnableModeration;
                    }
                    catch (ServerException ex)
                    {
                        // Handling the exception stating the "The object specified does not belong to a list."
                        if (ex.ServerErrorCode != -2113929210)
                        {
                            // TODO Replace this with an errorcode as well, does not work with localized o365 tenants
                            if (ex.Message.StartsWith("Cannot invoke method or retrieve property from null object. Object returned by the following call stack is null.") &&
                                ex.Message.Contains("ListItemAllFields"))
                            {
                                // E.g. custom display form aspx page being uploaded to the libraries Forms folder
                                normalFile = false;
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }
                }

                if (file.CheckOutType != CheckOutType.None || checkOutRequired)
                {
                    Log.Debug(Constants.LOGGING_SOURCE, "Checking in file '{0}'", file.Name);
                    file.CheckIn("Checked in by provisioning", publishingRequired ? CheckinType.MinorCheckIn : CheckinType.MajorCheckIn);
                    await context.ExecuteQueryRetryAsync();
                }

                if (level == FileLevel.Published)
                {
                    if (publishingRequired)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Publishing file '{0}'", file.Name);
                        file.Publish("Published by provisioning");
                        await context.ExecuteQueryRetryAsync();
                    }

                    if (approvalRequired)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Approving file '{0}'", file.Name);
                        file.Approve("Approved by provisioning");
                        await context.ExecuteQueryRetryAsync();
                    }
                }
            }
        }

        private static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern).
                               Replace(@"\*", ".*").
                               Replace(@"\?", ".") + "$";
        }

        /// <summary>
        /// Resets a file to its previous version.
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to process</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static async Task ResetFileToPreviousVersionAsync(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            await new SynchronizationContextRemover();
            await web.ResetFileToPreviousVersionImplementation(serverRelativeUrl, checkinType, comment);
        }

        /// <summary>
        /// Resets a file to its previous version.
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to process</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static void ResetFileToPreviousVersion(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            Task.Run(() => web.ResetFileToPreviousVersionImplementation(serverRelativeUrl, checkinType, comment)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Checks in a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkin</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static async Task ResetFileToPreviousVersionImplementation(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            var scope = new ConditionalScope(web.Context, () => !file.ServerObjectIsNull.Value && file.Exists && file.CheckOutType == CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file, f => f.Versions);
            }
            await web.Context.ExecuteQueryAsync();

            if (scope.TestResult.Value)
            {
                if (file.Versions.Count > 0)
                {
                    await web.CheckOutFileAsync(serverRelativeUrl);
                    var versionLabelPrevious = file.Versions[(file.Versions.Count - 1)].VersionLabel;
                    file.Versions.RestoreByLabel(versionLabelPrevious);
                }
                await web.CheckInFileAsync(serverRelativeUrl, checkinType, comment);
                await web.Context.ExecuteQueryAsync();
            }
        }
    }
}