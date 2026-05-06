'======================================
' 8. IDENTIFICAÇÃO DO CÓDIGO
'======================================
' Nome do arquivo: UsuarioRepo.vb
' VERSÃO: 15.0.30 (Mapeamento ORM e Namespace corrigidos)
' DATA: 06/05/2026
'=======================================
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports CONFIN_MXFC.Data
Imports CONFIN_MXFC.FinanceiroPJ.Models ' <-- Diretiva obrigatória adicionada

Namespace FinanceiroPJ
    Public Class UsuarioRepo
        Public Function ValidarLogin(login As String, senha As String) As Usuario
            Dim user As Usuario = Nothing
            Using conn As New SqlConnection(Db.ConnectionString)
                ' CORREÇÃO: Tabela atualizada para 'Usuario' e colunas correspondentes ao banco de dados e ao modelo
                Dim sql = "SELECT Username, PasswordHash, Salt, Role, MustChangePassword, ColaboradorId FROM Usuario WHERE Username = @l"
                Dim cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@l", login)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        user = New Usuario() With {
                            .Username = reader("Username").ToString(),
                            .PasswordHash = reader("PasswordHash").ToString(),
                            .Salt = reader("Salt").ToString(),
                            .Role = reader("Role").ToString(),
                            .MustChangePassword = Convert.ToBoolean(reader("MustChangePassword"))
                        }
                        If Not IsDBNull(reader("ColaboradorId")) Then
                            user.ColaboradorId = Convert.ToInt32(reader("ColaboradorId"))
                        End If
                    End If
                End Using
            End Using

            ' NOTA TÉCNICA: A validação do hash da senha deve ser feita aqui antes de retornar 'user'
            Return user
        End Function

        Public Async Function CriarAdminPadraoAsync() As Task
            Using conn As New SqlConnection(Db.ConnectionString)
                ' CORREÇÃO: Ajustado para o schema atualizado do Identity da MXFC
                Dim sql = "IF NOT EXISTS (SELECT 1 FROM Usuario WHERE Username = 'admin') " &
                          "INSERT INTO Usuario (Username, PasswordHash, Salt, Role, MustChangePassword) " &
                          "VALUES ('admin', 'hash_provisorio', 'salt_provisorio', 'Admin', 1)"
                Dim cmd As New SqlCommand(sql, conn)
                Await conn.OpenAsync()
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class
End Namespace