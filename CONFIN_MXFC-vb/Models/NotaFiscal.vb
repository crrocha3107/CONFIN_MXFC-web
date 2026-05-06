'======================================
' NotaFiscal.vb
' VERSÃO: 10.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 19/11/2025
'=======================================
Namespace FinanceiroPJ.Models
    Public Class NotaFiscal
        Public Property Id As Integer
        Public Property ColaboradorId As Integer
        Public Property Mes As Integer
        Public Property Ano As Integer
        Public Property ValorFixo As Decimal
        Public Property FG As Decimal
        Public Property Adicional As Decimal
        Public Property CentroCusto As String = ""
        Public Property Despesas As Decimal
        Public Property Contador As Decimal
        Public Property INSSProlabore As Decimal
        Public Property ImpostoNF As Decimal
        Public Property ValorNF As Decimal
        Public Property Prolabore As Decimal
        Public Property Observacao As String = ""
    End Class
End Namespace