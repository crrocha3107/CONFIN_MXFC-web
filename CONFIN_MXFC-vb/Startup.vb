'======================================
' Startup.vb
' VERSÃO: 15.0.40 (DI e Segurança)
'=======================================
Imports Microsoft.AspNetCore.Builder
Imports Microsoft.AspNetCore.Hosting
Imports Microsoft.Extensions.Configuration
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.AspNetCore.Authentication.JwtBearer
Imports Microsoft.IdentityModel.Tokens
Imports System.Text
Imports CONFIN_MXFC.Data

Namespace FinanceiroPJ
    Public Class Startup
        Public Property Configuration As IConfiguration

        Public Sub New(configuration As IConfiguration)
            Me.Configuration = configuration
        End Sub

        Public Sub ConfigureServices(services As IServiceCollection)
            services.AddControllers()

            ' Registro de Repositórios
            services.AddScoped(Of UsuarioRepo)()
            services.AddScoped(Of AuthService)()
            services.AddScoped(Of FuncionarioRepo)()

            ' JWT Config
            Dim jwtSecret = Configuration("Configuracoes:JwtSecret")
            Dim key = Encoding.ASCII.GetBytes(jwtSecret)

            services.AddAuthentication(Sub(auth)
                                           auth.DefaultAuthenticateScheme = "Bearer"
                                           auth.DefaultChallengeScheme = "Bearer"
                                       End Sub).AddJwtBearer(Sub(bearer)
                                                                 bearer.RequireHttpsMetadata = False
                                                                 bearer.SaveToken = True
                                                                 bearer.TokenValidationParameters = New TokenValidationParameters() With {
                                                                     .ValidateIssuerSigningKey = True,
                                                                     .IssuerSigningKey = New SymmetricSecurityKey(key),
                                                                     .ValidateIssuer = False,
                                                                     .ValidateAudience = False
                                                                 }
                                                             End Sub)

            services.AddCors(Sub(options)
                                 options.AddPolicy("AllowAll", Sub(builder)
                                                                   builder.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()
                                                               End Sub)
                             End Sub)
        End Sub

        Public Sub Configure(app As IApplicationBuilder, env As IWebHostEnvironment)
            app.UseRouting()
            app.UseCors("AllowAll")
            app.UseAuthentication()
            app.UseAuthorization()
            app.UseEndpoints(Sub(endpoints)
                                 endpoints.MapControllers()
                             End Sub)
        End Sub
    End Class
End Namespace