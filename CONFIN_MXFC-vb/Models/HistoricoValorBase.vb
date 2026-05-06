'======================================
' HistoricoValorBase.vb
' VERSÃO: 1.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 13/03/2026
'=======================================
Imports System

Namespace FinanceiroPJ.Models
    Public Class HistoricoValorBase
        Public Property Id As Integer
        Public Property FuncionarioREG As Integer
        Public Property DataAlteracao As Date
        Public Property ValorBase As Decimal
    End Class
End Namespace