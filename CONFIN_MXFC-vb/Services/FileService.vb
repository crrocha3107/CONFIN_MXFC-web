'======================================
' FileService.vb
' VERSÃO: 12.0.1 (by Rocha - MXFC Engenharia Ltda)
' DATA: 04/05/2026
'=======================================
Imports System.IO
Imports System
Imports System.Threading.Tasks

Namespace Services
    Public Class FileService
        Private ReadOnly _basePath As String

        ' Injeção do caminho base do servidor onde os anexos serão persistidos
        Public Sub New(basePath As String)
            _basePath = basePath
        End Sub

        Public Function GetFullPath(relativePath As String) As String
            Return Path.Combine(_basePath, "Anexos", Path.GetFileName(relativePath))
        End Function

        Public Function GetDeclaracoesPath() As String
            Dim pathDir = Path.Combine(_basePath, "Declaracoes")
            If Not Directory.Exists(pathDir) Then Directory.CreateDirectory(pathDir)
            Return pathDir
        End Function

        ' Método auxiliar para leitura assíncrona compatível com .NET Framework
        Private Async Function LerArquivoAsync(caminho As String) As Task(Of Byte())
            ' O último parâmetro 'True' habilita a I/O assíncrona no nível do SO
            Using fs As New FileStream(caminho, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, True)
                Using ms As New MemoryStream()
                    Await fs.CopyToAsync(ms)
                    Return ms.ToArray()
                End Using
            End Using
        End Function

        ' Operações assíncronas para ler arquivos sem bloquear a thread do servidor Web
        Public Async Function ObterAnexoBytesAsync(relativePath As String) As Task(Of Byte())
            If String.IsNullOrWhiteSpace(relativePath) Then Throw New Exception("O registro não possui um anexo.")
            Dim fullPath = GetFullPath(relativePath)
            If Not File.Exists(fullPath) Then Throw New FileNotFoundException($"O arquivo não foi encontrado no servidor:{vbCrLf}{fullPath}")
            Return Await LerArquivoAsync(fullPath)
        End Function

        Public Async Function ObterDeclaracaoBytesAsync(filename As String) As Task(Of Byte())
            If String.IsNullOrWhiteSpace(filename) Then Throw New Exception("O registro não possui declaração salva.")
            Dim fullPath = Path.Combine(GetDeclaracoesPath(), filename)
            If Not File.Exists(fullPath) Then Throw New FileNotFoundException($"O arquivo não foi encontrado no servidor:{vbCrLf}{fullPath}")
            Return Await LerArquivoAsync(fullPath)
        End Function

        ' Salva o stream de bytes recebido via upload HTTP
        Public Async Function SalvarAnexoAsync(stream As Stream, nomeOriginal As String) As Task(Of String)
            Dim anexoFolder = Path.Combine(_basePath, "Anexos")
            If Not Directory.Exists(anexoFolder) Then Directory.CreateDirectory(anexoFolder)
            Dim uniqueName = $"{Guid.NewGuid()}_{Path.GetFileName(nomeOriginal)}"
            Dim fullPath = Path.Combine(anexoFolder, uniqueName)

            Using fileStream As New FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, True)
                Await stream.CopyToAsync(fileStream)
            End Using
            Return uniqueName
        End Function

        Public Async Function SalvarDeclaracaoAsync(stream As Stream, nomeFuncionario As String) As Task(Of String)
            Dim declaracoesFolder = GetDeclaracoesPath()
            Dim safeName = String.Join("_", nomeFuncionario.Split(Path.GetInvalidFileNameChars()))
            Dim uniqueName = $"{safeName}_Declaracao_{Guid.NewGuid().ToString().Substring(0, 8)}.pdf"
            Dim fullPath = Path.Combine(declaracoesFolder, uniqueName)

            Using fileStream As New FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, True)
                Await stream.CopyToAsync(fileStream)
            End Using
            Return uniqueName
        End Function
    End Class
End Namespace