'======================================
' Db.vb
' VERSÃO: 15.0.1 (Migração Web)
'=======================================
Imports System.Data.SqlClient
Imports System.Threading.Tasks

Namespace Data
    Public Module Db
        ' A string agora deve ser alimentada pelo appsettings.json no Startup
        Public Property ConnectionString As String

        ' Método Assíncrono Otimizado para os Repositórios e chamadas Web
        Public Async Function GetConnectionAsync() As Task(Of SqlConnection)
            Dim cn = New SqlConnection(ConnectionString)
            Await cn.OpenAsync()
            Return cn
        End Function
    End Module
End Namespace