'======================================
' ContratoAditivo.vb
' VERSÃO: 2.0.0
' DATA: 03/12/2025
'=======================================
Namespace FinanceiroPJ.Models
    Public Class ContratoAditivo
        Public Property Id As Integer
        Public Property ContratoId As Integer
        Public Property Numero As String
        Public Property Data As Date
        Public Property DataFinal As Date ' Novo Prazo
        Public Property Valor As Decimal
        Public Property Descricao As String
        Public Property Anexo As Byte()
        Public Property TemAnexo As Boolean ' Auxiliar para a Grid
    End Class
End Namespace