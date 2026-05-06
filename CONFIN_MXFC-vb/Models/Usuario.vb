'======================================
' 8. IDENTIFICAÇÃO DO CÓDIGO
'======================================
' Nome do arquivo: Usuario.vb
' VERSÃO: 15.0.29 (Correção de mapeamento ORM)
' DATA: 06/05/2026
'=======================================
Namespace FinanceiroPJ.Models
    Public Class Usuario
        ' 1. Propriedades REAIS mapeadas exatamente como a tabela SQL (Migrations)
        Public Property Username As String
        Public Property PasswordHash As String
        Public Property Salt As String
        Public Property Role As String
        Public Property MustChangePassword As Boolean
        Public Property ColaboradorId As Integer?

        ' 2. Propriedades legadas mantidas em memória (não existem no DB atual)
        Public Property ID As Integer
        Public Property Nome As String
        Public Property Ativo As Boolean

        ' 3. Aliases de retrocompatibilidade para o código legado
        Public Property Login As String
            Get
                Return Username
            End Get
            Set(value As String)
                Username = value
            End Set
        End Property

        Public Property Perfil As String
            Get
                Return Role
            End Get
            Set(value As String)
                Role = value
            End Set
        End Property

        Public Property Senha As String
            Get
                Return PasswordHash
            End Get
            Set(value As String)
                PasswordHash = value
            End Set
        End Property
    End Class
End Namespace