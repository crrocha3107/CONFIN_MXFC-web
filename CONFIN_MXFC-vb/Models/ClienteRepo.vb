'======================================
' Cliente.vb
' VERSÃO: 1.0.0
' DATA: 02/12/2025
'=======================================
Namespace FinanceiroPJ.Models
    Public Class Cliente
        Public Property Id As Integer
        Public Property RazaoSocial As String
        Public Property CNPJ As String
        Public Property InscricaoEstadual As String
        Public Property Endereco As String
        Public Property CEP As String
        Public Property Telefone As String
        Public Property Responsavel As String
    End Class
End Namespace