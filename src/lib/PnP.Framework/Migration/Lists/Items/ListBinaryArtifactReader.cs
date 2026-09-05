using Microsoft.SharePoint.Client;
using PnP.Framework.Http;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.Json;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListBinaryArtifactReader
    {
        public static ListBinaryArtifactSnapshot Read(
            ClientContext context,
            Microsoft.SharePoint.Client.File file,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName,
            string fallbackServerRelativeUrl = null,
            string sourceArchiveStatus = null)
        {
            var serverRelativeUrl = string.IsNullOrWhiteSpace(fallbackServerRelativeUrl)
                ? file.ServerRelativeUrl
                : fallbackServerRelativeUrl;
            if (IsFullyArchived(sourceArchiveStatus))
            {
                var knownArchived = CaptureKnownArchived(
                    context,
                    serverRelativeUrl,
                    maximumBytes,
                    artifactStore,
                    mediaType,
                    originalName);
                if (knownArchived != null)
                {
                    return knownArchived;
                }
            }

            try
            {
                var streamResult = file.OpenBinaryStreamWithOptions(
                    SPOpenBinaryOptions.MinimizeProcessing);
                context.ExecuteQueryRetry();
                if (streamResult.Value == null)
                {
                    throw new FileNotFoundException("SharePoint returned no binary stream.");
                }
                using (streamResult.Value)
                {
                    return Capture(
                        streamResult.Value,
                        maximumBytes,
                        artifactStore,
                        mediaType,
                        originalName,
                        null);
                }
            }
            catch (Exception primaryException) when (IsDirectFallbackCandidate(primaryException))
            {
                return CaptureWithFallbacks(
                    context,
                    serverRelativeUrl,
                    maximumBytes,
                    artifactStore,
                    mediaType,
                    originalName,
                    primaryException);
            }
            catch (InvalidOperationException exception)
            {
                return new ListBinaryArtifactSnapshot
                {
                    Availability = EvidenceAvailability.Unavailable,
                    Diagnostics = { exception.Message }
                };
            }
        }

        internal static bool IsFullyArchived(string sourceArchiveStatus) =>
            string.Equals(
                sourceArchiveStatus?.Trim(),
                "fullyArchived",
                StringComparison.OrdinalIgnoreCase);

        private static ListBinaryArtifactSnapshot CaptureKnownArchived(
            ClientContext context,
            string serverRelativeUrl,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName)
        {
            var diagnostics = new List<string>
            {
                "KnownArchiveFastPath: source list item reports _FileArchiveStatus=fullyArchived; probing the authoritative REST content endpoint before CSOM stream variants."
            };
            try
            {
                var captured = CaptureViaRest(
                    context,
                    serverRelativeUrl,
                    maximumBytes,
                    artifactStore,
                    mediaType,
                    originalName,
                    "KnownArchiveFastPath: the REST content endpoint returned bytes, so the source archive hint was stale or reactivation completed during capture.");
                AddDiagnostics(captured, diagnostics);
                return captured;
            }
            catch (ArchivedContentException exception)
            {
                diagnostics.Add("KnownArchiveFastPath: " + exception.Message);
                return Unavailable(
                    diagnostics,
                    new[] { exception.Evidence });
            }
            catch (Exception exception) when (IsRetainableCaptureFailure(exception))
            {
                // The archive field is a routing hint, not authority. If the REST
                // response is inconclusive, retain the existing CSOM/fallback path.
                return null;
            }
        }

        private static ListBinaryArtifactSnapshot CaptureWithFallbacks(
            ClientContext context,
            string serverRelativeUrl,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName,
            Exception primaryException)
        {
            var diagnostics = new List<string>
            {
                "OpenBinaryStreamWithOptions(MinimizeProcessing): " + primaryException.Message
            };
            var archivedContentEvidence = new List<LiteralHttpArchivedContentEvidence>();
            var csomOptions = new[]
            {
                SPOpenBinaryOptions.None,
                SPOpenBinaryOptions.PreferNewStreamInterop,
                SPOpenBinaryOptions.MinimizeProcessing | SPOpenBinaryOptions.PreferNewStreamInterop,
                SPOpenBinaryOptions.MinimizeProcessing | SPOpenBinaryOptions.SkipVirusScan
            };
            foreach (var options in csomOptions)
            {
                var label = "OpenBinaryStreamWithOptions(" + options + ")";
                try
                {
                    var captured = CaptureViaCsomOptions(
                        context,
                        serverRelativeUrl,
                        options,
                        maximumBytes,
                        artifactStore,
                        mediaType,
                        originalName);
                    AddDiagnostics(captured, diagnostics);
                    captured.Diagnostics.Add(label + " supplied the retained payload.");
                    return captured;
                }
                catch (Exception exception) when (IsRetainableCaptureFailure(exception))
                {
                    diagnostics.Add(label + ": " + exception.Message);
                    if (IsMaximumSizeFailure(exception))
                    {
                        return Unavailable(diagnostics, archivedContentEvidence);
                    }
                }
            }

            try
            {
                var captured = CaptureViaRest(
                    context,
                    serverRelativeUrl,
                    maximumBytes,
                    artifactStore,
                    mediaType,
                    originalName,
                    "SharePoint REST file-content fallback supplied the retained payload.");
                AddDiagnostics(captured, diagnostics);
                return captured;
            }
            catch (ArchivedContentException exception)
            {
                archivedContentEvidence.Add(exception.Evidence);
                diagnostics.Add("REST file-content fallback: " + exception.Message);
                return Unavailable(diagnostics, archivedContentEvidence);
            }
            catch (Exception exception) when (IsRetainableCaptureFailure(exception))
            {
                diagnostics.Add("REST file-content fallback: " + exception.Message);
            }

            try
            {
                var captured = CaptureViaDownloadPage(
                    context,
                    serverRelativeUrl,
                    maximumBytes,
                    artifactStore,
                    mediaType,
                    originalName);
                AddDiagnostics(captured, diagnostics);
                captured.Diagnostics.Add("SharePoint download.aspx fallback supplied the retained payload.");
                return captured;
            }
            catch (ArchivedContentException exception)
            {
                archivedContentEvidence.Add(exception.Evidence);
                diagnostics.Add("download.aspx fallback: " + exception.Message);
            }
            catch (Exception exception) when (IsRetainableCaptureFailure(exception))
            {
                diagnostics.Add("download.aspx fallback: " + exception.Message);
            }

            return Unavailable(diagnostics, archivedContentEvidence);
        }

        private static ListBinaryArtifactSnapshot CaptureViaCsomOptions(
            ClientContext context,
            string serverRelativeUrl,
            SPOpenBinaryOptions options,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName)
        {
            using (var fallbackContext = context.Clone(context.Url))
            {
                var file = fallbackContext.Web.GetFileByServerRelativePath(
                    ResourcePath.FromDecodedUrl(Uri.UnescapeDataString(serverRelativeUrl)));
                var streamResult = file.OpenBinaryStreamWithOptions(options);
                fallbackContext.ExecuteQueryRetry();
                if (streamResult.Value == null)
                {
                    throw new FileNotFoundException("SharePoint returned no binary stream.");
                }

                using (streamResult.Value)
                {
                    return Capture(
                        streamResult.Value,
                        maximumBytes,
                        artifactStore,
                        mediaType,
                        originalName,
                        null);
                }
            }
        }

        private static ListBinaryArtifactSnapshot Capture(
            Stream source,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName,
            string diagnostic)
        {
            byte[] bytes;
            using (var buffer = new MemoryStream())
            {
                var block = new byte[81920];
                int read;
                while ((read = source.Read(block, 0, block.Length)) > 0)
                {
                    buffer.Write(block, 0, read);
                    if (buffer.Length > maximumBytes)
                    {
                        throw new InvalidOperationException("The list binary artifact exceeds the configured maximum dependency size.");
                    }
                }
                bytes = buffer.ToArray();
            }

            ArtifactReference reference;
            string contentBase64 = null;
            if (artifactStore == null)
            {
                reference = MigrationArtifact.Describe(bytes, mediaType, originalName);
                contentBase64 = Convert.ToBase64String(bytes);
            }
            else
            {
                using (var content = new MemoryStream(bytes, false))
                {
                    reference = artifactStore.Put(content, mediaType, originalName);
                }
            }

            var snapshot = new ListBinaryArtifactSnapshot
            {
                Artifact = reference,
                ContentBase64 = contentBase64,
                RepresentationKind = ListBinaryPayloadClassifier.Classify(bytes)
            };
            if (!string.IsNullOrWhiteSpace(diagnostic))
            {
                snapshot.Diagnostics.Add(diagnostic);
            }
            return snapshot;
        }

        private static ListBinaryArtifactSnapshot CaptureViaRest(
            ClientContext context,
            string serverRelativeUrl,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName,
            string successDiagnostic)
        {
            var oDataLiteral = "'" + (serverRelativeUrl ?? string.Empty).Replace("'", "''") + "'";
            var requestUri = context.Url.TrimEnd('/')
                + "/_api/web/GetFileByServerRelativePath(decodedurl=@path)/$value?@path="
                + Uri.EscapeDataString(oDataLiteral);
            using (var request = new HttpRequestMessage(HttpMethod.Get, requestUri))
            {
                PnPHttpClient.AuthenticateRequestAsync(request, context).GetAwaiter().GetResult();
                using (var client = PnPHttpClient.Instance.GetHttpClient(context))
                using (var response = client.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult())
                {
                    if (response.StatusCode == HttpStatusCode.Unauthorized
                        || response.StatusCode == HttpStatusCode.Forbidden)
                    {
                        response.EnsureSuccessStatusCode();
                    }
                    if (!response.IsSuccessStatusCode)
                    {
                        throw CreateHttpFailure(
                            response,
                            "list-binary-rest-value",
                            requestUri,
                            "SharePoint REST file-content request");
                    }
                    using (var content = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                    {
                        return Capture(
                            content,
                            maximumBytes,
                            artifactStore,
                            mediaType,
                            originalName,
                            successDiagnostic);
                    }
                }
            }
        }

        private static ListBinaryArtifactSnapshot CaptureViaDownloadPage(
            ClientContext context,
            string serverRelativeUrl,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName)
        {
            var requestUri = context.Url.TrimEnd('/')
                + "/_layouts/15/download.aspx?SourceUrl="
                + Uri.EscapeDataString(serverRelativeUrl ?? string.Empty);
            using (var request = new HttpRequestMessage(HttpMethod.Get, requestUri))
            {
                PnPHttpClient.AuthenticateRequestAsync(request, context).GetAwaiter().GetResult();
                using (var client = PnPHttpClient.Instance.GetHttpClient(context))
                using (var response = client.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult())
                {
                    if (response.StatusCode == HttpStatusCode.Unauthorized
                        || response.StatusCode == HttpStatusCode.Forbidden)
                    {
                        response.EnsureSuccessStatusCode();
                    }
                    if (!response.IsSuccessStatusCode)
                    {
                        throw CreateHttpFailure(
                            response,
                            "list-binary-download",
                            requestUri,
                            "SharePoint download request");
                    }

                    var responseMediaType = response.Content.Headers.ContentType?.MediaType;
                    if (!string.IsNullOrWhiteSpace(responseMediaType)
                        && responseMediaType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        throw new InvalidOperationException(
                            "SharePoint download request returned an HTML response instead of file bytes.");
                    }

                    using (var content = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                    {
                        return Capture(
                            content,
                            maximumBytes,
                            artifactStore,
                            mediaType,
                            originalName,
                            null);
                    }
                }
            }
        }

        private static void AddDiagnostics(
            ListBinaryArtifactSnapshot snapshot,
            IEnumerable<string> diagnostics)
        {
            foreach (var diagnostic in diagnostics)
            {
                snapshot.Diagnostics.Add(diagnostic);
            }
        }

        private static ListBinaryArtifactSnapshot Unavailable(
            IEnumerable<string> diagnostics,
            IList<LiteralHttpArchivedContentEvidence> archivedContentEvidence = null)
        {
            var snapshot = new ListBinaryArtifactSnapshot
            {
                Availability = EvidenceAvailability.Unavailable,
                ArchivedContentEvidence = archivedContentEvidence != null && archivedContentEvidence.Count > 0
                    ? archivedContentEvidence
                    : null
            };
            AddDiagnostics(snapshot, diagnostics);
            return snapshot;
        }

        private static Exception CreateHttpFailure(
            HttpResponseMessage response,
            string operation,
            string requestUri,
            string subject)
        {
            var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            string errorCode;
            string innerErrorCode;
            string message;
            if ((int)response.StatusCode == 423
                && TryReadError(body, out errorCode, out innerErrorCode, out message)
                && string.Equals(errorCode, "locked", StringComparison.OrdinalIgnoreCase)
                && string.Equals(innerErrorCode, "contentArchived", StringComparison.OrdinalIgnoreCase))
            {
                return new ArchivedContentException(
                    LiteralHttpArchivedContentEvidence.Create(
                        operation,
                        requestUri,
                        (int)response.StatusCode,
                        errorCode,
                        innerErrorCode,
                        message,
                        DateTimeOffset.UtcNow));
            }

            return new InvalidOperationException(
                subject
                + " returned HTTP "
                + (int)response.StatusCode
                + " ("
                + response.ReasonPhrase
                + ").");
        }

        private static bool TryReadError(
            string body,
            out string errorCode,
            out string innerErrorCode,
            out string message)
        {
            errorCode = null;
            innerErrorCode = null;
            message = null;
            if (string.IsNullOrWhiteSpace(body))
            {
                return false;
            }

            try
            {
                using (var document = JsonDocument.Parse(body))
                {
                    JsonElement error;
                    if (!document.RootElement.TryGetProperty("error", out error))
                    {
                        return false;
                    }

                    JsonElement value;
                    if (error.TryGetProperty("code", out value) && value.ValueKind == JsonValueKind.String)
                    {
                        errorCode = value.GetString();
                    }
                    if (error.TryGetProperty("message", out value) && value.ValueKind == JsonValueKind.String)
                    {
                        message = value.GetString();
                    }
                    JsonElement innerError;
                    if (error.TryGetProperty("innerError", out innerError)
                        && innerError.TryGetProperty("code", out value)
                        && value.ValueKind == JsonValueKind.String)
                    {
                        innerErrorCode = value.GetString();
                    }
                }
            }
            catch (JsonException)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(errorCode)
                && !string.IsNullOrWhiteSpace(innerErrorCode)
                && !string.IsNullOrWhiteSpace(message);
        }

        private static bool IsRetainableCaptureFailure(Exception exception) =>
            exception is ServerException
            || exception is IOException
            || exception is InvalidOperationException;

        private static bool IsMaximumSizeFailure(Exception exception) =>
            exception is InvalidOperationException
            && exception.Message.IndexOf(
                "exceeds the configured maximum dependency size",
                StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool IsDirectFallbackCandidate(Exception exception) =>
            exception is ServerException
            || exception is IOException
            || exception is InvalidOperationException
                && !IsMaximumSizeFailure(exception);

        private sealed class ArchivedContentException : InvalidOperationException
        {
            public ArchivedContentException(LiteralHttpArchivedContentEvidence evidence)
                : base("HTTP 423 locked/contentArchived: " + evidence?.Message)
            {
                Evidence = evidence ?? throw new ArgumentNullException(nameof(evidence));
            }

            public LiteralHttpArchivedContentEvidence Evidence { get; }
        }

        public static string MediaType(string path)
        {
            switch (Path.GetExtension(path ?? string.Empty).ToLowerInvariant())
            {
                case ".css": return "text/css";
                case ".js": return "application/javascript";
                case ".json": return "application/json";
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".svg": return "image/svg+xml";
                case ".xsl":
                case ".xslt": return "application/xml";
                case ".pdf": return "application/pdf";
                default: return "application/octet-stream";
            }
        }
    }
}
