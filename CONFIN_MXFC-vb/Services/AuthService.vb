'======================================
' 8. IDENTIFICAÇÃO DO CÓDIGO
'======================================
' Nome do arquivo: AuthService.vb
' VERSÃO: 15.0.31 (Correção de Namespace)
' DATA: 06/05/2026
'=======================================
Imports CONFIN_MXFC.FinanceiroPJ.Models ' <-- Diretiva obrigatória adicionada

Namespace FinanceiroPJ
    Public Class AuthService
        Private ReadOnly _repo As UsuarioRepo

        Public Sub New(repo As UsuarioRepo)
            _repo = repo
        End Sub

        Public Function Autenticar(usuario As String, senha As String) As Usuario
            Return _repo.ValidarLogin(usuario, senha)
        End Function
    End Class
End Namespace