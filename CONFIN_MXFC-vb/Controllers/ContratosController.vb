'======================================
' ContratosController.vb
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
    Public Class ContratosController
        Inherits ControllerBase

        Private ReadOnly _repo As ContratoRepo

        ' Injeção de dependência do repositório
        Public Sub New(repo As ContratoRepo)
            _repo = repo
        End Sub

        ' Rota: GET api/Contratos/cliente/{clienteId}
        <HttpGet("cliente/{clienteId}")>
        Public Async Function GetContratosPorCliente(clienteId As Integer) As Task(Of ActionResult(Of List(Of Contrato)))
            Dim lista = Await _repo.ListarAsync(clienteId)
            Return Ok(lista)
        End Function

        ' Rota: POST api/Contratos
        <HttpPost>
        Public Async Function PostContrato(<FromBody> c As Contrato) As Task(Of IActionResult)
            Await _repo.SalvarAsync(c)
            Return Ok(New With {.mensagem = "Contrato salvo com sucesso", .id = c.Id})
        End Function

        ' Rota: DELETE api/Contratos/{id}
        <HttpDelete("{id}")>
        Public Async Function DeleteContrato(id As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirAsync(id)
            Return NoContent()
        End Function

        ' Rota: GET api/Contratos/{contratoId}/anexos
        <HttpGet("{contratoId}/anexos")>
        Public Async Function GetAnexos(contratoId As Integer) As Task(Of ActionResult(Of List(Of ContratoAnexo)))
            Dim anexos = Await _repo.ListarAnexosAsync(contratoId)
            Return Ok(anexos)
        End Function

        ' Rota: DELETE api/Contratos/anexos/{anexoId}
        <HttpDelete("anexos/{anexoId}")>
        Public Async Function DeleteAnexo(anexoId As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirAnexoAsync(anexoId)
            Return NoContent()
        End Function
    End Class
End Namespace