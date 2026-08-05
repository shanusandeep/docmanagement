using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Scenario2.SharePointMaster.Components;
using Scenario2.SharePointMaster.Services;

var builder = WebApplication.CreateBuilder(args);

// Users sign in with Entra ID; the app checks their entitlement in its own
// database and projects it into SharePoint as just-in-time item permissions.
// The library custodian is the app-only service principal (SealedLibraryService).
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"))
    // auth-code flow (avoids needing the legacy id_token/implicit checkbox in Entra)
    .EnableTokenAcquisitionToCallDownstreamApi(["User.Read"])
    .AddInMemoryTokenCaches();

builder.Services.AddControllersWithViews().AddMicrosoftIdentityUI();
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection("SharePoint"));
builder.Services.Configure<DemoOptions>(builder.Configuration.GetSection("Demo"));

builder.Services.AddSingleton<ActivityLog>();
builder.Services.AddSingleton<WordTemplateService>();
builder.Services.AddSingleton<DocumentRegistry>();
builder.Services.AddSingleton<SealedLibraryService>();
builder.Services.AddHostedService<RevokeWorker>();

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

// Raw download of the current document.
app.MapGet("/api/doc/{itemId}/content",
    async (string itemId, string? name, SealedLibraryService library) =>
    {
        var stream = await library.GetContentAsync(itemId);
        if (stream == null) return Results.NotFound();
        return Results.Stream(stream,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name ?? $"{itemId}.docx");
    }).RequireAuthorization();

// Streams a specific version of a document (opens in Word as a .docx download).
app.MapGet("/api/doc/{itemId}/versions/{versionId}/content",
    async (string itemId, string versionId, string? name, SealedLibraryService library) =>
    {
        var stream = await library.GetVersionContentAsync(itemId, versionId);
        if (stream == null) return Results.NotFound();
        var fileName = $"{Path.GetFileNameWithoutExtension(name ?? itemId)}-v{versionId}.docx";
        return Results.Stream(stream,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
    }).RequireAuthorization();

// Current document converted to PDF by SharePoint.
app.MapGet("/api/doc/{itemId}/pdf",
    async (string itemId, string? name, SealedLibraryService library) =>
    {
        var stream = await library.GetPdfAsync(itemId);
        if (stream == null) return Results.NotFound();
        var fileName = $"{Path.GetFileNameWithoutExtension(name ?? itemId)}.pdf";
        return Results.Stream(stream, "application/pdf", fileName);
    }).RequireAuthorization();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
