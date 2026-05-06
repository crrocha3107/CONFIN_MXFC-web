'======================================
' AuthController.vb
' VERSÃO: 1.6.0 (Endpoint para Logout e limpeza de Cookie)
'=======================================
Imports Microsoft.AspNetCore.Mvc
Imports Microsoft.AspNetCore.Http
Imports System
Imports System.IdentityModel.Tokens.Jwt
Imports System.Security.Claims
Imports Microsoft.IdentityModel.Tokens
Imports System.Text
Imports Microsoft.Extensions.Configuration

Namespace FinanceiroPJ.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class AuthController
        Inherits ControllerBase

        Private ReadOnly _authService As AuthService
        Private ReadOnly _config As IConfiguration

        Public Sub New(authService As AuthService, config As IConfiguration)
            _authService = authService
            _config = config
        End Sub

        <HttpPost("login")>
        Public Function Login(<FromBody> request As LoginRequest) As IActionResult
            Try
                Dim usuarioLogado = _authService.Autenticar(request.Usuario, request.Senha)

                If usuarioLogado IsNot Nothing Then
                    Dim jwtSecret = _config("Configuracoes:JwtSecret")
                    Dim key = Encoding.ASCII.GetBytes(jwtSecret)

                    Dim tokenDescriptor = New SecurityTokenDescriptor With {
                        .Subject = New ClaimsIdentity(New Claim() {
                            New Claim(ClaimTypes.Name, request.Usuario)
                        }),
                        .Expires = DateTime.UtcNow.AddHours(8),
                        .SigningCredentials = New SigningCredentials(New SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
                    }

                    Dim tokenHandler = New JwtSecurityTokenHandler()
                    Dim tokenCriado = tokenHandler.CreateToken(tokenDescriptor)
                    Dim tokenString = tokenHandler.WriteToken(tokenCriado)

                    Return Ok(New With {Key .Token = tokenString})
                End If

                Return Unauthorized()
            Catch ex As Exception
                Return StatusCode(500, "Erro no servidor: " & ex.Message)
            End Try
        End Function

        <HttpPost("set-cookie")>
        Public Function SetCookie(<FromBody> token As String) As IActionResult
            Try
                Dim cookieOptions = New CookieOptions With {
                    .HttpOnly = True,
                    .Secure = False,
                    .SameSite = SameSiteMode.Strict,
                    .Expires = DateTime.UtcNow.AddHours(8)
                }
                Response.Cookies.Append("jwt_token", token, cookieOptions)
                Return Ok()
            Catch ex As Exception
                Return StatusCode(500, New With {Key .Erro = ex.Message})
            End Try
        End Function

        ' NOVO ENDPOINT DE LOGOUT PARA APAGAR O COOKIE DO NAVEGADOR
        <HttpPost("logout")>
        Public Function Logout() As IActionResult
            Try
                Response.Cookies.Delete("jwt_token")
                Return Ok()
            Catch ex As Exception
                Return StatusCode(500, New With {Key .Erro = ex.Message})
            End Try
        End Function
    End Class

    Public Class LoginRequest
        Public Property Usuario As String
        Public Property Senha As String
    End Class
End Namespace