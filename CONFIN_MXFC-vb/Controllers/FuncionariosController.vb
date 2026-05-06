'======================================
' FuncionariosController.vb
' VERSÃO: 2.2.0 (Busca e Listagem Otimizada)
'=======================================
Imports Microsoft.AspNetCore.Mvc
Imports CONFIN_MXFC.Data
Imports System.Threading.Tasks
Imports System.Linq

Namespace FinanceiroPJ.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class FuncionariosController
        Inherits ControllerBase

        Private ReadOnly _repo As FuncionarioRepo

        Public Sub New(repo As FuncionarioRepo)
            _repo = repo
        End Sub

        <HttpGet>
        Public Async Function GetFuncionarios() As Task(Of IActionResult)
            Try
                Dim lista = Await _repo.ListarAsync()
                Dim dadosGerais = lista.OrderBy(Function(f) f.NomeCompleto).
                                        Select(Function(f) New With {
                                            Key .REG = f.REG,
                                            Key .NomeCompleto = f.NomeCompleto,
                                            Key .CPF = f.CPF,
                                            Key .Cargo = f.Cargo,
                                            Key .DataAdmissao = f.DataAdmissao,
                                            Key .Telefone = f.Telefone,
                                            Key .Email = f.Email,
                                            Key .Ativo = f.Ativo
                                        }).ToList()
                Return Ok(dadosGerais)
            Catch ex As Exception
                Return StatusCode(500, "Erro interno: " & ex.Message)
            End Try
        End Function

        <HttpGet("{reg}")>
        Public Async Function GetFuncionario(reg As Integer) As Task(Of IActionResult)
            Try
                Dim lista = Await _repo.ListarAsync()
                Dim func = lista.FirstOrDefault(Function(f) f.REG = reg)
                If func Is Nothing Then Return NotFound()
                Return Ok(func)
            Catch ex As Exception
                Return StatusCode(500, "Erro: " & ex.Message)
            End Try
        End Function
    End Class
End Namespace