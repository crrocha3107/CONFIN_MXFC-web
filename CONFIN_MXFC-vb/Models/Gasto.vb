'======================================
' Gasto.vb
' VERSÃO: 10.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 19/11/2025
'=======================================
Namespace FinanceiroPJ.Models
    Public Class Gasto
        Public Property Id As Integer
        Public Property ColaboradorId As Integer
        Public Property [Data] As Date
        Public Property NomeDespesa As String = ""
        Public Property TipoDespesa As String = ""
        Public Property Valor As Decimal
        Public Property Anexo As Byte()
    End Class
End Namespace