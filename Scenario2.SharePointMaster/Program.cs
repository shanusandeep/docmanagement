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
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));

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
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
