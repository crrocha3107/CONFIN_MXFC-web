'======================================
' NotasFiscaisController.vb
' VERSÃO: 1.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports Microsoft.AspNetCore.Mvc
Imports CONFIN_MXFC.Data
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports CONFIN_MXFC.Services
Imports System.Threading.Tasks
Imports System.Collections.Generic

Namespace API.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class NotasFiscaisController
        Inherits ControllerBase

        Private ReadOnly _repo As NotaFiscalRepo
        Private ReadOnly _calcService As CalcService

        Public Sub New(repo As NotaFiscalRepo, calcService As CalcService)
            _repo = repo
            _calcService = calcService
        End Sub

        ' Rota: GET api/NotasFiscais/obter?colabId=1&mes=5&ano=2026
        <HttpGet("obter")>
        Public Async Function GetNotaFiscal(<FromQuery> colabId As Integer, <FromQuery> mes As Integer, <FromQuery> ano As Integer) As Task(Of ActionResult(Of NotaFiscal))
            Dim nf = Await _repo.ObterOuCriarAsync(colabId, mes, ano)
            Return Ok(nf)
        End Function

        ' Rota: POST api/NotasFiscais
        <HttpPost>
        Public Async Function PostNotaFiscal(<FromBody> nf As NotaFiscal) As Task(Of IActionResult)
            ' Recalcula no backend para garantir a integridade dos dados antes de salvar no SQL (Security Best Practice)
            _calcService.CalcularNotaFiscal(nf)
            Await _repo.SalvarAsync(nf)
            Return Ok(New With {.mensagem = "Nota Fiscal processada e salva com sucesso", .nota = nf})
        End Function
    End Class
End Namespace