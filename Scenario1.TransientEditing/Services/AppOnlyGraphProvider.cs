using Azure.Identity;
using Microsoft.Graph;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// App-only Graph client (client credentials) used by the background sweep,
/// which runs without a signed-in user. Interactive actions stay delegated.
/// </summary>
public class AppOnlyGraphProvider
{
    public GraphServiceClient? Client { get; }
    public bool IsConfigured => Client != null;

    public AppOnlyGraphProvider(IConfiguration config)
    {
        var tenantId = config["AzureAd:TenantId"];
        var clientId = config["AzureAd:ClientId"];
        var clientSecret = config["AzureAd:ClientSecret"];

        if (string.IsNullOrWhiteSpace(tenantId) || tenantId.StartsWith("YOUR_") ||
            string.IsNullOrWhiteSpace(clientId) || clientId.StartsWith("YOUR_") ||
            string.IsNullOrWhiteSpace(clientSecret) || clientSecret.StartsWith("set-via"))
        {
            Client = null;
            return;
        }

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        Client = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
    }
}
