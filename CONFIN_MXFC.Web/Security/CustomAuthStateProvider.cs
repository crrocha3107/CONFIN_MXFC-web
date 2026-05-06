/*
======================================
Arquivo: CustomAuthStateProvider.cs
VERSÃO: 1.5.0 (Modernização C# 12 - IDE0290, IDE0090)
MXFC Engenharia Ltda.
======================================
*/
using Microsoft.AspNetCore.Components.Authorization;
using System.Security.Claims;
using System.IdentityModel.Tokens.Jwt;

namespace CONFIN_MXFC.Web.Security;

// IDE0290: Uso de Construtor Primário direto na definição da classe
public class CustomAuthStateProvider(IHttpContextAccessor httpContextAccessor) : AuthenticationStateProvider
{
    // IDE0090: Expressão 'new' simplificada
    private readonly ClaimsPrincipal _anonymous = new(new ClaimsIdentity());

    public override Task<AuthenticationState> GetAuthenticationStateAsync()
    {
        try
        {
            var token = httpContextAccessor.HttpContext?.Request.Cookies["jwt_token"];

            if (string.IsNullOrWhiteSpace(token))
                return Task.FromResult(new AuthenticationState(_anonymous));

            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadJwtToken(token);

            if (jwtToken.ValidTo < DateTime.UtcNow)
                return Task.FromResult(new AuthenticationState(_anonymous));

            var identity = new ClaimsIdentity(jwtToken.Claims, "jwt");
            var user = new ClaimsPrincipal(identity);

            return Task.FromResult(new AuthenticationState(user));
        }
        catch
        {
            return Task.FromResult(new AuthenticationState(_anonymous));
        }
    }

    public async Task MarcarUsuarioComoAutenticado(string _)
    {
        // IDE0060: Parâmetro 'token' alterado para '_' por não ser usado no corpo
        NotifyAuthenticationStateChanged(GetAuthenticationStateAsync());
        await Task.CompletedTask;
    }

    public async Task MarcarUsuarioComoDeslogado()
    {
        NotifyAuthenticationStateChanged(Task.FromResult(new AuthenticationState(_anonymous)));
        await Task.CompletedTask;
    }
}