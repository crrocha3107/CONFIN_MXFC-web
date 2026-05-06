'======================================
' HistoricoValorBaseController.vb
' VERSÃO: 1.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports Microsoft.AspNetCore.Mvc
Imports CONFIN_MXFC.Data
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System

Namespace API.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class HistoricoValorBaseController
        Inherits ControllerBase

        Private ReadOnly _repo As HistoricoValorBaseRepo

        Public Sub New(repo As HistoricoValorBaseRepo)
            _repo = repo
        End Sub

        ' Rota: GET api/HistoricoValorBase/funcionario/{reg}
        <HttpGet("funcionario/{reg}")>
        Public Async Function GetHistorico(reg As Integer) As Task(Of ActionResult(Of List(Of HistoricoValorBase)))
            Dim lista = Await _repo.ListarAsync(reg)
            Return Ok(lista)
        End Function

        ' Rota: POST api/HistoricoValorBase
        <HttpPost>
        Public Async Function PostHistorico(<FromBody> h As HistoricoValorBase) As Task(Of IActionResult)
            Await _repo.InserirHistoricoAsync(h.FuncionarioREG, h.DataAlteracao, h.ValorBase)
            Return Ok(New With {.mensagem = "Histórico inserido com sucesso"})
        End Function

        ' Rota: PUT api/HistoricoValorBase/{id}
        <HttpPut("{id}")>
        Public Async Function PutHistorico(id As Integer, <FromBody> h As HistoricoValorBase) As Task(Of IActionResult)
            Await _repo.AtualizarHistoricoAsync(id, h.DataAlteracao, h.ValorBase)
            Return NoContent()
        End Function

        ' Rota: DELETE api/HistoricoValorBase/{id}
        <HttpDelete("{id}")>
        Public Async Function DeleteHistorico(id As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirHistoricoAsync(id)
            Return NoContent()
        End Function
    End Class
End Namespace