Namespace FinanceiroPJ.Models
    Public Class Colaborador
        Public Property Id As Integer
        Public Property NomeEmpresa As String = ""
        Public Property CNPJ As String = ""
        Public Property Endereco As String = ""
        Public Property Cidade As String = ""
        Public Property CEP As String = ""
        Public Property Responsavel As String = ""
        Public Property Telefone As String = ""
        Public Property Email As String = ""
        Public Property PixOuDadosBancarios As String = ""

        ' O REG da Ficha Cadastral (PF) vinculada a este Colaborador (PJ)
        Public Property FuncionarioREG As Integer?

        ' --- CORREÇÃO ADICIONADA ---
        ' Isso diz ao ComboBox para usar o NomeEmpresa como o texto de exibição.
        Public Overrides Function ToString() As String
            Return NomeEmpresa
        End Function
        ' --- FIM DA CORREÇÃO ---
    End Class
End Namespace