using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Scenario1.TransientEditing.Components;
using Scenario1.TransientEditing.Services;

var builder = WebApplication.CreateBuilder(args);

string[] graphScopes = ["User.Read", "Files.ReadWrite.All"];

// Delegated auth: the app always acts as the signed-in user, so SharePoint's
// "Modified by" and Track Changes attribution show the real person.
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi(graphScopes)
    .AddMicrosoftGraph("https://graph.microsoft.com/v1.0", graphScopes)
    .AddInMemoryTokenCaches();

builder.Services.AddControllersWithViews().AddMicrosoftIdentityUI();
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddMicrosoftIdentityConsentHandler();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection("SharePoint"));
builder.Services.Configure<DemoOptions>(builder.Configuration.GetSection("Demo"));

builder.Services.AddSingleton<ActivityLog>();
builder.Services.AddSingleton<WordTemplateService>();
builder.Services.AddSingleton<NetworkShareService>();
builder.Services.AddSingleton<DocumentStateStore>();
builder.Services.AddSingleton<AppOnlyGraphProvider>();
builder.Services.AddSingleton<SyncBackEngine>();
builder.Services.AddScoped<TransientDocumentService>();
builder.Services.AddHostedService<SyncBackWorker>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

app.MapControllers();

// PDF conversion (delegated): in-place for checked-out docs, transient upload for share-resident docs
app.MapGet("/api/pdf", async (string path, TransientDocumentService docs) =>
{
    var (content, fileName) = await docs.ConvertToPdfAsync(path);
    return Results.Stream(content, "application/pdf", fileName);
}).RequireAuthorization();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
