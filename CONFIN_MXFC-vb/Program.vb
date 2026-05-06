'======================================
' Program.vb
' VERSÃO: 15.0.4 (Correção de Inicialização e Tratamento DB)
'=======================================
Imports System
Imports System.IO
Imports Microsoft.AspNetCore.Hosting
Imports Microsoft.Extensions.Hosting
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.Configuration
Imports CONFIN_MXFC.Data

Namespace FinanceiroPJ
    Module Program
        Sub Main(args As String())
            ' 1. Força o diretório de trabalho para a pasta onde os arquivos foram publicados
            Dim caminhoExecucao = AppContext.BaseDirectory
            Directory.SetCurrentDirectory(caminhoExecucao)

            ' 2. Cria o Host (Renomeado para appHost para evitar erro BC30980)
            Dim appHost As IHost = Host.CreateDefaultBuilder(args) _
                .ConfigureAppConfiguration(Sub(configContext, configBuilder)
                                               configBuilder.SetBasePath(caminhoExecucao)
                                               configBuilder.AddJsonFile("appsettings.json", optional:=False, reloadOnChange:=True)
                                           End Sub) _
                .ConfigureWebHostDefaults(Sub(webBuilder)
                                              webBuilder.UseStartup(Of Startup)()
                                              webBuilder.UseUrls("http://localhost:5000")
                                          End Sub) _
                .Build()

            ' 3. Puxa a configuração de forma segura
            Dim config = appHost.Services.GetRequiredService(Of IConfiguration)()
            Db.ConnectionString = config.GetConnectionString("DefaultConnection")

            ' 4. Validação Técnica
            If String.IsNullOrWhiteSpace(Db.ConnectionString) Then
                Dim msg = $"ERRO: String 'DefaultConnection' não encontrada." &
                          $"Verifique se o arquivo existe em: {caminhoExecucao}appsettings.json"
                Console.WriteLine(msg)
            Else
                ' 5. Executa as migrações com Proteção de Exceção (Banco: 172.25.25.12)
                Try
                    Console.WriteLine("Tentando conectar ao Banco de Dados SQL Server...")
                    Migrations.EnsureCreatedAsync().GetAwaiter().GetResult()
                    Console.WriteLine("Conexao e Migracoes concluidas com sucesso.")
                Catch ex As Exception
                    Console.WriteLine("==================================================")
                    Console.WriteLine("FALHA CRÍTICA DE REDE/BANCO DE DADOS (IP 172.25.25.12)")
                    Console.WriteLine($"ERRO: {ex.Message}")
                    Console.WriteLine("O Servidor Web será iniciado na porta 5000, mas o login falhará.")
                    Console.WriteLine("==================================================")
                End Try
            End If

            ' 6. Inicia o servidor Kestrel independentemente do status do banco
            appHost.Run()
        End Sub

        Public Function CreateHostBuilder(args As String()) As IHostBuilder
            Return Host.CreateDefaultBuilder(args) _
                .ConfigureWebHostDefaults(Sub(webBuilder)
                                              webBuilder.UseStartup(Of Startup)()
                                              webBuilder.UseUrls("http://localhost:5000")
                                          End Sub)
        End Function
    End Module
End Namespace