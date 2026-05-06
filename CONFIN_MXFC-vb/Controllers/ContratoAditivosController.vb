'======================================
' ContratoAditivosController.vb
' VERSÃO: 1.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports CONFIN_MXFC.Data
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports Microsoft.AspNetCore.Mvc

Namespace API.Controllers
    <ApiController>
    <Route("api/[controller]")>
    Public Class ContratoAditivosController
        Inherits ControllerBase

        Private ReadOnly _repo As ContratoAditivoRepo

        ' Injeção de dependência do repositório
        Public Sub New(repo As ContratoAditivoRepo)
            _repo = repo
        End Sub

        ' Rota: GET api/ContratoAditivos/contrato/{contratoId}
        <HttpGet("contrato/{contratoId}")>
        Public Async Function GetAditivosPorContrato(contratoId As Integer) As Task(Of ActionResult(Of List(Of ContratoAditivo)))
            Dim lista = Await _repo.ListarAsync(contratoId)
            Return Ok(lista)
        End Function

        ' Rota: POST api/ContratoAditivos
        <HttpPost>
        Public Async Function PostAditivo(<FromBody> a As ContratoAditivo) As Task(Of IActionResult)
            Await _repo.SalvarAsync(a)
            Return Ok(New With {.mensagem = "Aditivo salvo com sucesso"})
        End Function

        ' Rota: DELETE api/ContratoAditivos/{id}
        <HttpDelete("{id}")>
        Public Async Function DeleteAditivo(id As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirAsync(id)
            Return NoContent()
        End Function
    End Class
End Namespace