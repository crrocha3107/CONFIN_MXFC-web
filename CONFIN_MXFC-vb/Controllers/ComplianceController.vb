'======================================
' ComplianceController.vb
' VERSÃO: 1.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 06/05/2026
' DESCRIÇÃO: Endpoint para gerenciar a emissão e recebimento de PDFs (BLOB).
'=======================================
Imports System.IO
Imports System.Threading.Tasks
Imports CONFIN_MXFC.Data
Imports Microsoft.AspNetCore.Http
Imports Microsoft.AspNetCore.Mvc

Namespace API.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class ComplianceController
        Inherits ControllerBase

        Private ReadOnly _repo As FuncionarioRepo

        Public Sub New(repo As FuncionarioRepo)
            _repo = repo
        End Sub

        ' Rota: POST api/Compliance/gerar-termo/{reg}
        <HttpPost("gerar-termo/{reg}")>
        Public Async Function GerarTermo(reg As Integer) As Task(Of IActionResult)
            Try
                Dim func = Await _repo.ObterAsync(reg)
                If func Is Nothing Then Return NotFound("Colaborador não encontrado.")

                ' TODO: Aqui você deve integrar com o seu ComplianceService.vb (iTextSharp)
                ' que gera o documento com a folha timbrada da MXFC.
                ' Simulação de retorno de bytes:
                ' Dim pdfBytes As Byte() = ComplianceService.GerarTermoCompliance(func)
                ' Return File(pdfBytes, "application/pdf", $"Termo_Compliance_{reg}.pdf")

                Return Ok() ' Remova este Ok() quando plugar o gerador real
            Catch ex As Exception
                Return StatusCode(500, "Erro ao gerar termo: " & ex.Message)
            End Try
        End Function

        ' Rota: PUT api/Compliance/importar-termo/{reg}
        <HttpPut("importar-termo/{reg}")>
        Public Async Function ImportarTermo(reg As Integer, file As IFormFile) As Task(Of IActionResult)
            Try
                If file Is Nothing OrElse file.Length = 0 Then
                    Return BadRequest("Nenhum arquivo válido foi recebido.")
                End If

                Using ms As New MemoryStream()
                    Await file.CopyToAsync(ms)
                    Dim pdfBytes = ms.ToArray()

                    ' Salva o BLOB no banco de dados via repositório
                    Dim func = Await _repo.ObterAsync(reg)
                    If func IsNot Nothing Then
                        func.Declaracao = pdfBytes
                        ' Atualiza o funcionário no banco
                        Await _repo.SalvarAsync(func)
                        Return Ok()
                    Else
                        Return NotFound("Colaborador não localizado.")
                    End If
                End Using
            Catch ex As Exception
                Return StatusCode(500, "Erro ao arquivar PDF: " & ex.Message)
            End Try
        End Function

        ' Rota: GET api/Compliance/ver-termo/{reg}
        <HttpGet("ver-termo/{reg}")>
        Public Async Function VerTermo(reg As Integer) As Task(Of IActionResult)
            Try
                Dim func = Await _repo.ObterAsync(reg)

                ' Verifica se o funcionário existe e se a coluna Declaracao (BLOB) contém dados
                If func Is Nothing OrElse func.Declaracao Is Nothing OrElse func.Declaracao.Length = 0 Then
                    Return NotFound("Declaração não encontrada.")
                End If

                ' Retorna os bytes como um arquivo PDF real para o navegador fazer o download
                Return File(func.Declaracao, "application/pdf", $"Declaracao_{reg}.pdf")
            Catch ex As Exception
                Return StatusCode(500, "Erro ao ler PDF do banco: " & ex.Message)
            End Try
        End Function
    End Class
End Namespace