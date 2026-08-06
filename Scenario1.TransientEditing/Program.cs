using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Scenario1.TransientEditing.Components;
using Scenario1.TransientEditing.Services;

var builder = WebApplication.CreateBuilder(args);

// Sites.Selected (delegated): acts as the user, but only on sites explicitly
// granted to this app — no tenant-wide file access is ever requested.
string[] graphScopes = ["User.Read", "Sites.Selected"];

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

// Raw download: SharePoint copy when checked out, share copy otherwise
app.MapGet("/api/download", async (string path, TransientDocumentService docs) =>
{
    var (content, fileName) = await docs.DownloadAsync(path);
    return Results.Stream(content, MimeFor(fileName), fileName);
}).RequireAuthorization();

static string MimeFor(string name) => Path.GetExtension(name).ToLowerInvariant() switch
{
    ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pdf" => "application/pdf",
    ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    _ => "application/octet-stream"
};

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
