'======================================
' Funcionario.vb
' VERSÃO: 11.0.0 (by Rocha - MXFC Engenharia Ltda)
' DATA: 21/11/2025
'=======================================
Namespace FinanceiroPJ.Models
    Public Class Funcionario
        Public Property REG As Integer
        Public Property NomeCompleto As String = ""
        Public Property Foto As Byte()

        Public Property Declaracao As Byte()

        Public Property Situacao As String = "Ativo"
        Public Property DataNascimento As Date?
        Public Property Naturalidade As String = ""
        Public Property Nacionalidade As String = ""
        Public Property EstadoCivil As String = ""
        Public Property NomeConjuge As String = ""
        Public Property NomePai As String = ""
        Public Property NomeMae As String = ""
        Public Property Endereco As String = ""
        Public Property EnderecoNumero As String = ""
        Public Property EnderecoComplemento As String = ""
        Public Property EnderecoBairro As String = ""
        Public Property CEP As String = ""
        Public Property Cidade As String = ""
        Public Property Estado As String = ""
        Public Property Telefone As String = ""
        Public Property Celular As String = ""
        Public Property Email As String = ""
        Public Property Email2 As String = ""
        Public Property Formacao As String = ""
        Public Property Cargo As String = ""
        Public Property DataAdmissao As Date?
        Public Property RegimeContratacao As String = ""
        Public Property Salario As Decimal = 0
        Public Property PrazoExperiencia As String = ""
        Public Property HorarioTrabalho As String = ""
        Public Property DescontaVT As String = ""
        Public Property ObservacoesRH As String = ""
        Public Property RG As String = ""
        Public Property RGDataExpedicao As Date?
        Public Property RGOrgaoExp As String = ""
        Public Property RGUF As String = ""
        Public Property CPF As String = ""
        Public Property PIS As String = ""
        Public Property CTPS As String = ""
        Public Property CTPSSerie As String = ""
        Public Property CTPSUF As String = ""
        Public Property CTPSDataExp As Date?
        Public Property TituloEleitor As String = ""
        Public Property TituloZona As String = ""
        Public Property TituloSecao As String = ""
        Public Property TituloMunicipioUF As String = ""
        Public Property TituloDataEmissao As Date?
        Public Property Reservista As String = ""
        Public Property CNH As String = ""
        Public Property CNHCategoria As String = ""
        Public Property CNHValidade As Date?
        Public Property CNHData1Hab As Date?
        Public Property CNHDataEmissao As Date?
        Public Property ConjugeDataNasc As Date?
        Public Property ConjugeRG As String = ""
        Public Property ConjugeRGUF As String = ""
        Public Property ConjugeCPF As String = ""
        Public Property Banco As String = ""
        Public Property Agencia As String = ""
        Public Property ContaCorrente As String = ""
        Public Property ContaModalidade As String = ""
        Public Property PIX As String = ""
        Public Property Banco2 As String = ""
        Public Property Agencia2 As String = ""
        Public Property ContaCorrente2 As String = ""
        Public Property ContaModalidade2 As String = ""
        Public Property PIX2 As String = ""
        Public Property Dependentes As String = ""
        Public Property EmergenciaNome As String = ""
        Public Property EmergenciaFone As String = ""
        Public Property EmergenciaParentesco As String = ""
        Public Property Ativo As Boolean = True

        Public Overrides Function ToString() As String
            Return NomeCompleto
        End Function
    End Class
End Namespace