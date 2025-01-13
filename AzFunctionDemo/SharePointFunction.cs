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

            // Required auth details from the environment settings
            var clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
            var clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET");
            var tenantId = Environment.GetEnvironmentVariable("TENANT_ID");

            // Create a Graph client to access data from our sharepoint site
            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
            var graphClient = new GraphServiceClient(clientSecretCredential);

            // Required site details from the environment settings to access data in a list. Will be sent with the request
            var siteId = Environment.GetEnvironmentVariable("SITE_ID");
            var listId = Environment.GetEnvironmentVariable("LIST_ID");

            // Retrieve siteId and listId from query parameters
            //var siteId = req.Query["siteId"];
            //var listId = req.Query["listId"];

            //This is just a check the site connectivity. I will retun this value if there is issue with connection to lists
            var lists = await graphClient.Sites[siteId].GetAsync();

            //Query list items from graph
            var listItems = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync(config =>
            {
                config.QueryParameters.Expand = new string[] { "fields($select=Title,Document,Author)" };
            });

            var selectedFields = listItems.Value?.Select(item =>
            {
                var fields = item.Fields.AdditionalData;
                return new
                {
                    Title = fields.ContainsKey("Title") ? fields["Title"] : null,
                    Document = fields.ContainsKey("Document") ? fields["Document"] : null,
                    Author = fields.ContainsKey("Author") ? fields["Author"] : null
                };
            });

            return new OkObjectResult(selectedFields);

        }
    }
}
