'======================================
' ColaboradoresController.vb
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
    Public Class ColaboradoresController
        Inherits ControllerBase

        Private ReadOnly _repo As ColaboradorRepo

        Public Sub New(repo As ColaboradorRepo)
            _repo = repo
        End Sub

        ' Rota: GET api/Colaboradores
        <HttpGet>
        Public Async Function GetColaboradores() As Task(Of ActionResult(Of List(Of Colaborador)))
            Dim lista = Await _repo.ListarAsync()
            Return Ok(lista)
        End Function

        ' Rota: POST api/Colaboradores (Usado para Insert e Update baseado no Id do objeto)
        <HttpPost>
        Public Async Function PostColaborador(<FromBody> c As Colaborador) As Task(Of ActionResult(Of Integer))
            Dim id = Await _repo.SalvarAsync(c)
            Return Ok(id)
        End Function

        ' Rota: DELETE api/Colaboradores/{id}
        <HttpDelete("{id}")>
        Public Async Function DeleteColaborador(id As Integer) As Task(Of IActionResult)
            Await _repo.ExcluirAsync(id)
            Return NoContent()
        End Function
    End Class
End Namespace