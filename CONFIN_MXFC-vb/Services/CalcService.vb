'======================================
' CalcService.vb
' VERSÃO: 9.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports System

Namespace Services
    Public Class CalcService
        Private Const Pct_Imposto As Double = 0.06
        Private Const Pct_INSS As Double = 0.11
        Private Const Pct_Prolabore As Double = 0.28

        Public Function CalcularProlaboreFlex(salarioLiquido As Decimal, custoContador As Decimal) As Decimal
            Dim denominador = 1 - Pct_Imposto - (Pct_Prolabore * Pct_INSS)
            If Math.Abs(denominador) < 0.0000001 Then Return 0D
            Dim valorNF = (salarioLiquido + custoContador) / CDec(denominador)
            Return valorNF * CDec(Pct_Prolabore)
        End Function

        Public Sub CalcularNotaFiscal(ByRef nf As FinanceiroPJ.Models.NotaFiscal)
            ' 1. Somar todos os custos base, INCLUINDO O ADICIONAL
            Dim custosBase = nf.ValorFixo + nf.FG + nf.Adicional + nf.Despesas + nf.Contador

            ' 2. Calcular o denominador
            Dim denominador = 1 - CDec(Pct_Imposto) - (CDec(Pct_Prolabore) * CDec(Pct_INSS))

            ' 3. Evitar divisão por zero
            If Math.Abs(denominador) < 0.0000001 Then denominador = 1D

            ' 4. Calcular os valores
            nf.ValorNF = custosBase / denominador
            nf.Prolabore = nf.ValorNF * CDec(Pct_Prolabore)
            nf.INSSProlabore = nf.Prolabore * CDec(Pct_INSS)
            nf.ImpostoNF = nf.ValorNF * CDec(Pct_Imposto)
        End Sub
    End Class
End Namespace