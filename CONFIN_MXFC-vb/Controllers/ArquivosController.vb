'======================================
' ArquivosController.vb
' VERSÃO: 1.0.1 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports Microsoft.AspNetCore.Mvc
Imports Microsoft.AspNetCore.Http
Imports CONFIN_MXFC.Services
Imports System.Threading.Tasks
Imports System

Namespace API.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class ArquivosController
        Inherits ControllerBase

        Private ReadOnly _fileService As FileService

        ' Injeção de dependência via construtor, seguindo as melhores práticas do .NET 8
        Public Sub New(fileService As FileService)
            _fileService = fileService
        End Sub

        ' Rota: GET api/Arquivos/anexo/nome_do_arquivo.ext
        ' CORREÇÃO: Sintaxe de parâmetro de rota ajustada para VB.NET
        <HttpGet("anexo/{nomeArquivo}")>
        Public Async Function DownloadAnexo(nomeArquivo As String) As Task(Of IActionResult)
            Try
                If String.IsNullOrEmpty(nomeArquivo) Then Return BadRequest(New With {.mensagem = "Nome do arquivo não fornecido."})

                Dim bytes = Await _fileService.ObterAnexoBytesAsync(nomeArquivo)
                ' Retorna o arquivo para download forçado no navegador
                Return File(bytes, "application/octet-stream", nomeArquivo)
            Catch ex As Exception
                Return NotFound(New With {.mensagem = ex.Message})
            End Try
        End Function

        ' Rota: GET api/Arquivos/declaracao/nome_da_declaracao.pdf
        ' CORREÇÃO: Sintaxe de parâmetro de rota ajustada para VB.NET
        <HttpGet("declaracao/{nomeArquivo}")>
        Public Async Function DownloadDeclaracao(nomeArquivo As String) As Task(Of IActionResult)
            Try
                If String.IsNullOrEmpty(nomeArquivo) Then Return BadRequest(New With {.mensagem = "Nome da declaração não fornecido."})

                Dim bytes = Await _fileService.ObterDeclaracaoBytesAsync(nomeArquivo)
                ' Retorna o PDF para visualização direta no navegador (Content-Disposition: inline)
                Return File(bytes, "application/pdf")
            Catch ex As Exception
                Return NotFound(New With {.mensagem = ex.Message})
            End Try
        End Function

        ' Rota: POST api/Arquivos/upload-anexo
        <HttpPost("upload-anexo")>
        Public Async Function UploadAnexo(arquivo As IFormFile) As Task(Of IActionResult)
            ' Verificação de nulidade conforme documentação oficial do IFormFile
            If arquivo Is Nothing OrElse arquivo.Length = 0 Then
                Return BadRequest(New With {.mensagem = "Nenhum arquivo recebido pelo servidor."})
            End If

            Try
                ' Utilizando OpenReadStream para processamento eficiente de memória
                Using stream = arquivo.OpenReadStream()
                    Dim uniqueName = Await _fileService.SalvarAnexoAsync(stream, arquivo.FileName)
                    Return Ok(New With {.nomeArquivoSalvo = uniqueName})
                End Using
            Catch ex As Exception
                Return StatusCode(500, New With {.mensagem = String.Format("Erro interno de processamento: {0}", ex.Message)})
            End Try
        End Function
    End Class


End Namespace