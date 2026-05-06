'======================================
' AppSettings.vb
' VERSÃO: 15.0.0 (Migração Web)
'=======================================
Namespace App
    Public Module AppSettings
        ' Define a pasta de dados dentro da raiz da aplicação no servidor
        Public Property DbFolder As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data")

        Public ReadOnly Property AnexosFolder As String
            Get
                Dim path = IO.Path.Combine(DbFolder, "Anexos")
                If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
                Return path
            End Get
        End Property
    End Module
End Namespace