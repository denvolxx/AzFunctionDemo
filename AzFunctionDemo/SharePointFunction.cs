using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace AzFunctionDemo
{
    public class SharePointFunction(ILogger<SharePointFunction> logger)
    {

        [Function("GetSharePointData")]
        public async Task<IActionResult> RunAsync([HttpTrigger(AuthorizationLevel.Function, "get", Route = "get-sharepoint-data")] HttpRequest req)
        {
            logger.LogInformation("C# HTTP trigger function processed a request.");

            // Authentication details
            var clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
            var clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET");
            var tenantId = Environment.GetEnvironmentVariable("TENANT_ID");

            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

            // Create a Graph client
            var graphClient = new GraphServiceClient(clientSecretCredential);

            var siteId = Environment.GetEnvironmentVariable("SITE_ID");
            var listId = Environment.GetEnvironmentVariable("LIST_ID");

            //Access Denied here
            //var listItems = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync(config =>
            //{
            //    config.QueryParameters.Expand = new string[] { "fields($select=Name,Color,Quantity)" };
            //});

            var lists = await graphClient.Sites[siteId].GetAsync();

            return new OkObjectResult(lists);
        }
    }
}
