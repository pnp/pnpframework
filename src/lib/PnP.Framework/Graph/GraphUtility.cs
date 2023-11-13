﻿using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions.Authentication;
using PnP.Framework.Diagnostics;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Graph
{
    public class TokenProvider : IAccessTokenProvider
    {
        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
            CancellationToken cancellationToken = default)
        {
            var token = AccessToken;
            // get the token and return it in your own way
            return Task.FromResult(token);
        }
        public string AccessToken { get; set; }
        public AllowedHostsValidator AllowedHostsValidator { get; }
    }

    /// <summary>
    /// Utility class to perform Graph operations.
    /// </summary>
    public static class GraphUtility
    {
        private const int defaultRetryCount = 10;
        private const int defaultDelay = 500;

        /// <summary>
        ///  Creates a new GraphServiceClient instance using a custom PnPHttpProvider
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to configure the HTTP bearer Authorization Header</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud deployment to use.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph</param>
        /// <returns></returns>
#pragma warning disable CA2000
        public static GraphServiceClient CreateGraphClient(string accessToken, int retryCount = defaultRetryCount, int delay = defaultDelay, AzureEnvironment azureEnvironment = AzureEnvironment.Production, bool useBetaEndPoint = false)
        {
            var baseUrl = $"https://{AuthenticationManager.GetGraphEndPoint(azureEnvironment)}/{(useBetaEndPoint ? "beta" : "v1.0")}";
            // Creates a new GraphServiceClient instance using a custom PnPHttpProvider
            // which natively supports retry logic for throttled requests
            // Default are 10 retries with a base delay of 500ms
            //var result = new GraphServiceClient(baseUrl, new DelegateAuthenticationProvider(
            //            async (requestMessage) =>
            //            {
            //                await Task.Run(() =>
            //                {
            //                    if (!string.IsNullOrEmpty(accessToken))
            //                    {
            //                        // Configure the HTTP bearer Authorization Header
            //                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
            //                    }
            //                });
            //            }), new PnPHttpProvider(retryCount, delay));

            var tokenProvider = new TokenProvider();
            tokenProvider.AccessToken = accessToken;

            var authenticationProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
            var result = new GraphServiceClient(authenticationProvider, baseUrl);

            return (result);
        }
#pragma warning restore CA2000

        /// <summary>
        /// This method sends an Azure guest user invitation to the provided email address.
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="guestUserEmail">Email of the user to whom the invite must be sent</param>
        /// <param name="redirectUri">URL where the user will be redirected after the invite is accepted.</param>
        /// <param name="customizedMessage">Customized email message to be sent in the invitation email.</param>
        /// <param name="guestUserDisplayName">Display name of the Guest user.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        public static Invitation InviteGuestUser(string accessToken, string guestUserEmail, string redirectUri, string customizedMessage = "", string guestUserDisplayName = "", AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            Invitation inviteUserResponse = null;

            try
            {
                Invitation invite = new Invitation
                {
                    InvitedUserEmailAddress = guestUserEmail
                };
                if (!string.IsNullOrWhiteSpace(guestUserDisplayName))
                {
                    invite.InvitedUserDisplayName = guestUserDisplayName;
                }
                invite.InviteRedirectUrl = redirectUri;
                invite.SendInvitationMessage = true;

                // Form the invite email message body
                if (!string.IsNullOrWhiteSpace(customizedMessage))
                {
                    InvitedUserMessageInfo inviteMsgInfo = new InvitedUserMessageInfo
                    {
                        CustomizedMessageBody = customizedMessage
                    };
                    invite.InvitedUserMessageInfo = inviteMsgInfo;
                }

                // Create the graph client and send the invitation.
                GraphServiceClient graphClient = CreateGraphClient(accessToken, azureEnvironment: azureEnvironment);
                inviteUserResponse = graphClient.Invitations.PostAsync(invite).Result;
            }
            catch (ODataError ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return inviteUserResponse;
        }
    }
}
