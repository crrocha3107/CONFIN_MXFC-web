'======================================
' GastosController.vb
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
    Public Class GastosController
        Inherits ControllerBase

        Private ReadOnly _repo As GastoRepo

        ' Injeção de dependência do repositório
        Public Sub New(repo As GastoRepo)
            _repo = repo
        End Sub

        ' Rota: GET api/Gastos/listar?colabId=1&mes=5&ano=2026
        <HttpGet("listar")>
        Public Async Function GetGastos(<FromQuery> colabId As Integer, <FromQuery> mes As Integer?, <FromQuery> ano As Integer?) As Task(Of ActionResult(Of List(Of Gasto)))
            Dim lista = Await _repo.ListarAsync(colabId, mes, ano)
            Return Ok(lista)
        End Function

        ' Rota: GET api/Gastos/somar?colabId=1&mes=5&ano=2026
        <HttpGet("somar")>
        Public Async Function GetSomaGastos(<FromQuery> colabId As Integer, <FromQuery> mes As Integer, <FromQuery> ano As Integer) As Task(Of ActionResult(Of Decimal))
            Dim soma = Await _repo.SomarDoPeriodoAsync(colabId, mes, ano)
            Return Ok(soma)
        End Function

        ' Rota: GET api/Gastos/nomes-despesa
        <HttpGet("nomes-despesa")>
        Public Async Function GetNomesDespesa() As Task(Of ActionResult(Of List(Of String)))
            Dim nomes = Await _repo.ListarNomesDespesaAsync()
            Return Ok(nomes)
        End Function

        ' Rota: GET api/Gastos/periodos-disponiveis
        <HttpGet("periodos-disponiveis")>
        Public Async Function GetPeriodosDisponiveis() As Task(Of ActionResult(Of List(Of Periodo)))
            Dim periodos = Await _repo.ListarPeriodosDisponiveisAsync()
            Return Ok(periodos)
        End Function

        ' Rota: POST api/Gastos
        <HttpPost>
        Public Async Function PostGasto(<FromBody> g As Gasto) As Task(Of ActionResult(Of Integer))
            Dim id = Await _repo.SalvarAsync(g)
            Return Ok(id)
        End Function

        ' Rota: DELETE api/Gastos/{id}
        <HttpDelete("{id}")>
        Public Async Function DeleteGasto(id As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirAsync(id)
            Return NoContent()
        End Function
    End Class
End Namespace