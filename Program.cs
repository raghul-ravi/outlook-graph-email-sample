using Azure.Identity;
using Microsoft.Graph;

var builder = WebApplication.CreateBuilder(args);

// Build GraphServiceClient with client credentials from environment variables
var tenantId = builder.Configuration["AZURE_TENANT_ID"];
var clientId = builder.Configuration["AZURE_CLIENT_ID"];
var clientSecret = builder.Configuration["AZURE_CLIENT_SECRET"];

if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
{
    throw new InvalidOperationException("Azure AD credentials are not configured.");
}

var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
var graphClient = new GraphServiceClient(credential);

builder.Services.AddSingleton(graphClient);

builder.Services.AddControllers();

var app = builder.Build();

app.MapControllers();

app.Run();
