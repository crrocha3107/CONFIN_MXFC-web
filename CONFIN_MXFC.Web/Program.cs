/*
======================================
Arquivo: Program.cs
VERSÃO: 1.4.0 (Correção de Delegates e Fluxo de Build)
MXFC Engenharia Ltda.
======================================
*/
using CONFIN_MXFC.Web.Components;
using CONFIN_MXFC.Web.Security;
using Microsoft.AspNetCore.Components.Authorization;

var builder = WebApplication.CreateBuilder(args);

// --- 1. CONFIGURAÇÃO DE SERVIÇOS ---
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddHttpContextAccessor();

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri("http://localhost:5000") });

builder.Services.AddAuthorizationCore();
builder.Services.AddScoped<CustomAuthStateProvider>();
builder.Services.AddScoped<AuthenticationStateProvider>(p => p.GetRequiredService<CustomAuthStateProvider>());
builder.Services.AddCascadingAuthenticationState();

// --- 2. CONSTRUÇÃO DO APP ---
var app = builder.Build();

// --- 3. MIDDLEWARES ---
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseStaticFiles();
app.UseAntiforgery();

// --- 4. MAPEAMENTO DE ROTAS (Local Auth) ---
app.MapPost("/auth/local-login", async (HttpContext context, [Microsoft.AspNetCore.Mvc.FromBody] string token) =>
{
    var cookieOptions = new CookieOptions
    {
        HttpOnly = true,
        Secure = false,
        SameSite = SameSiteMode.Lax,
        Expires = DateTime.UtcNow.AddHours(8)
    };

    context.Response.Cookies.Append("jwt_token", token, cookieOptions);
    return await Task.FromResult(Results.Ok());
});

app.MapPost("/auth/local-logout", (HttpContext context) =>
{
    context.Response.Cookies.Delete("jwt_token");
    return Results.Ok();
});

// --- 5. COMPONENTES ---
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();