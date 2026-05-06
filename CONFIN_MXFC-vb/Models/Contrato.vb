'======================================
' Contrato.vb
' VERSÃO: 2.0.0
' DATA: 03/12/2025
'=======================================
Imports System.Collections.Generic

Namespace FinanceiroPJ.Models
    Public Class Contrato
        Public Property Id As Integer
        Public Property ClienteId As Integer
        Public Property NumeroContrato As String
        Public Property Valor As Decimal
        Public Property DataInicial As Date
        Public Property DataFinal As Date
        Public Property Prazo As Integer ' Em dias
        Public Property CentroCusto As String

        Public Property ValorAditivos As Decimal = 0
        Public ReadOnly Property ValorTotal As Decimal
            Get
                Return Valor + ValorAditivos
            End Get
        End Property

        Public Property Anexos As New List(Of ContratoAnexo)
    End Class
End Namespace