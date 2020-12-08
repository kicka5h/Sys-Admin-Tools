using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;

namespace Graph_Utilities
{
    //https://www.c-sharpcorner.com/article/how-to-get-domain-users-search-users-and-user-from-active-directory-using-net/
    public class Accounts
    {
        private static GraphServiceClient _graphServiceClient;

        private static GraphServiceClient GetAuthenticatedGraphClient()
        {
            var authenticationProvider = CreateAuthorizationProvider();
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static IAuthenticationProvider CreateAuthorizationProvider()
        {
            var clientId = System.Environment.GetEnvironmentVariable("AzureADAppClientId", EnvironmentVariableTarget.Process);
            var clientSecret = System.Environment.GetEnvironmentVariable("AzureADAppClientSecret", EnvironmentVariableTarget.Process);
            var redirectUri = System.Environment.GetEnvironmentVariable("AzureADAppRedirectUri", EnvironmentVariableTarget.Process);
            var tenantId = System.Environment.GetEnvironmentVariable("AzureADAppTenantId", EnvironmentVariableTarget.Process);
            var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

            //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithRedirectUri(redirectUri)
                                              .WithClientSecret(clientSecret)
                                              .Build();

            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        public static List<User> GetAllUsers()
        {
            //Query using Graph SDK (preferred when possible)
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient();

            List<User> AllUsers = new List<User>();
            String Properties = "displayName";
            IGraphServiceUsersCollectionPage users = graphServiceClient.Users.Request().Select(Properties).GetAsync().Result;
            bool QueryIncomplete = false;
            do
            {
                QueryIncomplete = false;
                AllUsers.AddRange(users);
                if (users.NextPageRequest != null)
                {
                    users = users.NextPageRequest.GetAsync().Result;
                    QueryIncomplete = true;
                }

            } while (QueryIncomplete);

            return AllUsers;
        }
    }
}
