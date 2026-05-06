'======================================
' ExcelService.vb
' VERSÃO: 15.0.0 (Exclusivo Web - Streams)
'=======================================
Imports ClosedXML.Excel
Imports CONFIN_MXFC.Data
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports ExcelDataReader
Imports System
Imports System.Data

Namespace Services
    Public Module ExcelService
        Private ReadOnly culture As New CultureInfo("pt-BR")

        Private ReadOnly headersFuncionario As String() = {
            "REG", "NomeCompleto", "Situacao", "DataNascimento", "Naturalidade", "Nacionalidade", "EstadoCivil", "NomeConjuge", "NomePai", "NomeMae",
            "Endereco", "EnderecoNumero", "EnderecoComplemento", "EnderecoBairro", "CEP", "Cidade", "Estado", "Telefone", "Celular", "Email", "Email2", "Formacao",
            "Cargo", "DataAdmissao", "RegimeContratacao", "Salario", "PrazoExperiencia", "HorarioTrabalho", "DescontaVT", "ObservacoesRH",
            "RG", "RGDataExpedicao", "RGOrgaoExp", "RGUF", "CPF", "PIS", "CTPS", "CTPSSerie", "CTPSUF", "CTPSDataExp",
            "TituloEleitor", "TituloZona", "TituloSecao", "TituloMunicipioUF", "TituloDataEmissao", "Reservista", "CNH", "CNHCategoria", "CNHValidade", "CNHData1Hab", "CNHDataEmissao",
            "ConjugeDataNasc", "ConjugeRG", "ConjugeRGUF", "ConjugeCPF",
            "Banco", "Agencia", "ContaCorrente", "ContaModalidade", "PIX",
            "Banco2", "Agencia2", "ContaCorrente2", "ContaModalidade2", "PIX2",
            "Dependentes", "EmergenciaNome", "EmergenciaFone", "EmergenciaParentesco",
            "Ativo"
         }

        Private Sub SetCellDate(cell As IXLCell, dt As Date?)
            If dt.HasValue Then
                cell.Value = dt.Value
                cell.Style.DateFormat.Format = "dd/MM/yyyy"
            Else
                cell.Value = ""
            End If
        End Sub

        Private Sub PreencherPlanilhaColaboradores(wb As XLWorkbook, itens As IEnumerable(Of Colaborador))
            Dim ws = wb.AddWorksheet("Colaboradores")
            Dim headers = {"NomeEmpresa", "CNPJ", "Endereco", "Cidade", "CEP", "Responsavel", "Telefone", "Email", "PixOuDadosBancarios"}
            For i = 0 To headers.Length - 1
                ws.Cell(1, i + 1).Value = headers(i)
            Next
            Dim r = 2
            For Each c In itens
                ws.Cell(r, 1).Value = c.NomeEmpresa
                ws.Cell(r, 2).Value = c.CNPJ
                ws.Cell(r, 3).Value = c.Endereco
                ws.Cell(r, 4).Value = c.Cidade
                ws.Cell(r, 5).Value = c.CEP
                ws.Cell(r, 6).Value = c.Responsavel
                ws.Cell(r, 7).Value = c.Telefone
                ws.Cell(r, 8).Value = c.Email
                ws.Cell(r, 9).Value = c.PixOuDadosBancarios
                r += 1
            Next
        End Sub

        Public Function ExportarColaboradoresBytes(itens As IEnumerable(Of Colaborador)) As Byte()
            Using wb As New XLWorkbook()
                PreencherPlanilhaColaboradores(wb, itens)
                Using ms As New MemoryStream()
                    wb.SaveAs(ms)
                    Return ms.ToArray()
                End Using
            End Using
        End Function

        Private Sub PreencherPlanilhaRelatorio(wb As XLWorkbook, itens As List(Of Object))
            Dim ws = wb.AddWorksheet("Resumo")
            ws.Cell(1, 1).Value = "Colaborador"
            ws.Cell(1, 2).Value = "Mês"
            ws.Cell(1, 3).Value = "Ano"
            ws.Cell(1, 4).Value = "ValorNotaFiscal"
            ws.Row(1).Style.Font.Bold = True
            Dim r = 2
            For Each it In itens
                ws.Cell(r, 1).Value = CStr(it.GetType().GetProperty("Colaborador").GetValue(it))
                ws.Cell(r, 2).Value = CInt(it.GetType().GetProperty("Mes").GetValue(it))
                ws.Cell(r, 3).Value = CInt(it.GetType().GetProperty("Ano").GetValue(it))
                ws.Cell(r, 4).Value = CDec(it.GetType().GetProperty("ValorNotaFiscal").GetValue(it))
                r += 1
            Next
            ws.Column(4).Style.NumberFormat.Format = "R$ #,##0.00"
            ws.Columns().AdjustToContents()
        End Sub

        Public Function ExportarRelatorioBytes(itens As List(Of Object)) As Byte()
            Using wb As New XLWorkbook()
                PreencherPlanilhaRelatorio(wb, itens)
                Using ms As New MemoryStream()
                    wb.SaveAs(ms)
                    Return ms.ToArray()
                End Using
            End Using
        End Function

        Private Sub PreencherPlanilhaFuncionarios(wb As XLWorkbook, itens As IEnumerable(Of Funcionario))
            Dim ws = wb.AddWorksheet("FichasCadastrais")
            For i = 0 To headersFuncionario.Length - 1
                ws.Cell(1, i + 1).Value = headersFuncionario(i)
                ws.Cell(1, i + 1).Style.Font.Bold = True
            Next
            Dim r = 2
            For Each f In itens
                ws.Cell(r, 1).Value = f.REG
                ws.Cell(r, 2).Value = f.NomeCompleto
                ws.Cell(r, 3).Value = f.Situacao
                SetCellDate(ws.Cell(r, 4), f.DataNascimento)
                ws.Cell(r, 5).Value = f.Naturalidade
                ws.Cell(r, 6).Value = f.Nacionalidade
                ws.Cell(r, 7).Value = f.EstadoCivil
                ws.Cell(r, 8).Value = f.NomeConjuge
                ws.Cell(r, 9).Value = f.NomePai
                ws.Cell(r, 10).Value = f.NomeMae
                ws.Cell(r, 11).Value = f.Endereco
                ws.Cell(r, 12).Value = f.EnderecoNumero
                ws.Cell(r, 13).Value = f.EnderecoComplemento
                ws.Cell(r, 14).Value = f.EnderecoBairro
                ws.Cell(r, 15).Value = f.CEP
                ws.Cell(r, 16).Value = f.Cidade
                ws.Cell(r, 17).Value = f.Estado
                ws.Cell(r, 18).Value = f.Telefone
                ws.Cell(r, 19).Value = f.Celular
                ws.Cell(r, 20).Value = f.Email
                ws.Cell(r, 21).Value = f.Email2
                ws.Cell(r, 22).Value = f.Formacao
                ws.Cell(r, 23).Value = f.Cargo
                SetCellDate(ws.Cell(r, 24), f.DataAdmissao)
                ws.Cell(r, 25).Value = f.RegimeContratacao
                ws.Cell(r, 26).Value = f.Salario
                ws.Cell(r, 27).Value = f.PrazoExperiencia
                ws.Cell(r, 28).Value = f.HorarioTrabalho
                ws.Cell(r, 29).Value = f.DescontaVT
                ws.Cell(r, 30).Value = f.ObservacoesRH
                ws.Cell(r, 31).Value = f.RG
                SetCellDate(ws.Cell(r, 32), f.RGDataExpedicao)
                ws.Cell(r, 33).Value = f.RGOrgaoExp
                ws.Cell(r, 34).Value = f.RGUF
                ws.Cell(r, 35).Value = f.CPF
                ws.Cell(r, 36).Value = f.PIS
                ws.Cell(r, 37).Value = f.CTPS
                ws.Cell(r, 38).Value = f.CTPSSerie
                ws.Cell(r, 39).Value = f.CTPSUF
                SetCellDate(ws.Cell(r, 40), f.CTPSDataExp)
                ws.Cell(r, 41).Value = f.TituloEleitor
                ws.Cell(r, 42).Value = f.TituloZona
                ws.Cell(r, 43).Value = f.TituloSecao
                ws.Cell(r, 44).Value = f.TituloMunicipioUF
                SetCellDate(ws.Cell(r, 45), f.TituloDataEmissao)
                ws.Cell(r, 46).Value = f.Reservista
                ws.Cell(r, 47).Value = f.CNH
                ws.Cell(r, 48).Value = f.CNHCategoria
                SetCellDate(ws.Cell(r, 49), f.CNHValidade)
                SetCellDate(ws.Cell(r, 50), f.CNHData1Hab)
                SetCellDate(ws.Cell(r, 51), f.CNHDataEmissao)
                SetCellDate(ws.Cell(r, 52), f.ConjugeDataNasc)
                ws.Cell(r, 53).Value = f.ConjugeRG
                ws.Cell(r, 54).Value = f.ConjugeRGUF
                ws.Cell(r, 55).Value = f.ConjugeCPF
                ws.Cell(r, 56).Value = f.Banco
                ws.Cell(r, 57).Value = f.Agencia
                ws.Cell(r, 58).Value = f.ContaCorrente
                ws.Cell(r, 59).Value = f.ContaModalidade
                ws.Cell(r, 60).Value = f.PIX
                ws.Cell(r, 61).Value = f.Banco2
                ws.Cell(r, 62).Value = f.Agencia2
                ws.Cell(r, 63).Value = f.ContaCorrente2
                ws.Cell(r, 64).Value = f.ContaModalidade2
                ws.Cell(r, 65).Value = f.PIX2
                ws.Cell(r, 66).Value = f.Dependentes
                ws.Cell(r, 67).Value = f.EmergenciaNome
                ws.Cell(r, 68).Value = f.EmergenciaFone
                ws.Cell(r, 69).Value = f.EmergenciaParentesco
                ws.Cell(r, 70).Value = f.Ativo
                r += 1
            Next
            ws.Column(26).Style.NumberFormat.Format = "R$ #,##0.00"
            ws.Columns().AdjustToContents()
        End Sub

        Public Function ExportarFuncionariosBytes(itens As IEnumerable(Of Funcionario)) As Byte()
            Using wb As New XLWorkbook()
                PreencherPlanilhaFuncionarios(wb, itens)
                Using ms As New MemoryStream()
                    wb.SaveAs(ms)
                    Return ms.ToArray()
                End Using
            End Using
        End Function

        Private Function GetDataSet(stream As Stream) As DataSet
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
            Using reader = ExcelReaderFactory.CreateReader(stream)
                Return reader.AsDataSet(New ExcelDataSetConfiguration() With {
                    .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                        .UseHeaderRow = False
                    }
                })
            End Using
        End Function

        Public Function ImportarColaboradoresStream(stream As Stream) As Integer
            Return ProcessarImportacaoColaboradores(GetDataSet(stream))
        End Function

        Private Function ProcessarImportacaoColaboradores(ds As DataSet) As Integer
            Dim repo As New ColaboradorRepo()
            Dim count As Integer = 0
            Dim dt = ds.Tables(0)
            For r As Integer = 1 To dt.Rows.Count - 1
                Dim row = dt.Rows(r)
                If row.Item(0) Is DBNull.Value OrElse String.IsNullOrWhiteSpace(row.Item(0).ToString()) Then
                    Continue For
                End If
                Dim c As New Colaborador With {
                    .NomeEmpresa = GetVal(dt, r, 0),
                    .CNPJ = GetVal(dt, r, 1),
                    .Endereco = GetVal(dt, r, 2),
                    .Cidade = GetVal(dt, r, 3),
                    .CEP = GetVal(dt, r, 4),
                    .Responsavel = GetVal(dt, r, 5),
                    .Telefone = GetVal(dt, r, 6),
                    .Email = GetVal(dt, r, 7),
                    .PixOuDadosBancarios = GetVal(dt, r, 8)
                }
                ' Salva utilizando a versão síncrona/adaptada do repositório para o fluxo de importação
                repo.SalvarAsync(c).GetAwaiter().GetResult()
                count += 1
            Next
            Return count
        End Function

        Public Function ImportarFuncionariosStream(stream As Stream, Optional reutilizarReg As Boolean = False) As Integer
            Return ProcessarImportacaoFuncionarios(GetDataSet(stream), reutilizarReg)
        End Function

        Private Function ProcessarImportacaoFuncionarios(ds As DataSet, reutilizarReg As Boolean) As Integer
            Dim repo As New FuncionarioRepo()
            Dim count As Integer = 0
            If ds.Tables.Count = 0 Then Throw New Exception("O arquivo Excel não contém planilhas.")
            Dim dt = ds.Tables(0)

            ' 1. Pré-calcular variáveis com lógica condicional ANTES de criar o objeto Funcionario
            Dim telFixo = (GetVal(dt, "D15") & " " & GetVal(dt, "E15")).Trim()

            Dim celPart1 = (GetVal(dt, "G15") & " " & GetVal(dt, "H15") & " " & GetVal(dt, "I15")).Trim()
            Dim celPart2 = (GetVal(dt, "L15") & " " & GetVal(dt, "M15")).Trim()

            Dim celularFinal As String
            If Not String.IsNullOrEmpty(celPart2) Then
                celularFinal = (celPart1 & " / " & celPart2).Trim()
            Else
                celularFinal = celPart1
            End If

            ' 2. Inicialização simplificada do objeto (Exigência IDE0017 resolvida)
            ' Note o uso do ponto (.) e a separação por vírgulas (,)
            Dim f As New Funcionario() With {
                .NomeCompleto = GetVal(dt, "D8").ToUpper(),
                .Situacao = "Ativo",
                .Endereco = GetVal(dt, "D9"),
                .EnderecoNumero = GetVal(dt, "Q9"),
                .EnderecoComplemento = GetVal(dt, "D10"),
                .EnderecoBairro = GetVal(dt, "H10"),
                .CEP = GetVal(dt, "O10"),
                .Cidade = GetVal(dt, "D11"),
                .Estado = GetVal(dt, "L11"),
                .Nacionalidade = GetVal(dt, "D12"),
                .DataNascimento = GetDateVal(dt, "I12"),
                .NomeMae = GetVal(dt, "D13"),
                .NomePai = GetVal(dt, "D14"),
                .Telefone = telFixo,
                .Celular = celularFinal,
                .Email = GetVal(dt, "D16"),
                .Email2 = GetVal(dt, "L16"),
                .Formacao = ParseEscolaridade(dt),
                .EstadoCivil = ParseEstadoCivil(dt),
                .NomeConjuge = GetVal(dt, "E20"),
                .ConjugeDataNasc = GetDateVal(dt, "P20"),
                .ConjugeRG = GetVal(dt, "E21"),
                .ConjugeRGUF = GetVal(dt, "H21"),
                .ConjugeCPF = GetVal(dt, "J21"),
                .CTPS = GetVal(dt, "C24"),
                .CTPSSerie = GetVal(dt, "G24"),
                .CTPSUF = GetVal(dt, "J24"),
                .CTPSDataExp = GetDateVal(dt, "M24"),
                .PIS = GetVal(dt, "P24"),
                .RG = GetVal(dt, "C25"),
                .RGUF = GetVal(dt, "G25"),
                .RGOrgaoExp = GetVal(dt, "J25"),
                .RGDataExpedicao = GetDateVal(dt, "O25"),
                .CPF = GetVal(dt, "C26"),
                .Naturalidade = GetVal(dt, "I26"),
                .TituloEleitor = GetVal(dt, "D27"),
                .TituloZona = GetVal(dt, "G27"),
                .TituloSecao = GetVal(dt, "I27"),
                .TituloMunicipioUF = GetVal(dt, "L27"),
                .TituloDataEmissao = GetDateVal(dt, "P27"),
                .CNH = GetVal(dt, "E28"),
                .CNHCategoria = GetVal(dt, "H28"),
                .CNHValidade = GetDateVal(dt, "J28"),
                .CNHData1Hab = GetDateVal(dt, "M28"),
                .CNHDataEmissao = GetDateVal(dt, "P28"),
                .Banco = GetVal(dt, "C31"),
                .Agencia = GetVal(dt, "E31"),
                .ContaCorrente = GetVal(dt, "H31"),
                .ContaModalidade = GetVal(dt, "L31"),
                .PIX = GetVal(dt, "N31"),
                .Banco2 = GetVal(dt, "C32"),
                .Agencia2 = GetVal(dt, "E32"),
                .ContaCorrente2 = GetVal(dt, "H32"),
                .ContaModalidade2 = GetVal(dt, "L32"),
                .PIX2 = GetVal(dt, "N32"),
                .DataAdmissao = GetDateVal(dt, "E45"),
                .PrazoExperiencia = GetVal(dt, "E46"),
                .Cargo = GetVal(dt, "H45"),
                .Salario = GetDecimalVal(dt, "E47"),
                .HorarioTrabalho = GetVal(dt, "E48"),
                .ObservacoesRH = GetVal(dt, "H47"),
                .DescontaVT = ParseVT(dt)
            }

            ' 3. Salva de forma adaptada para o fluxo
            repo.SalvarAsync(f, reutilizarReg).GetAwaiter().GetResult()
            count += 1

            Return count
        End Function

        Private Sub ParseAddress(addr As String, ByRef rowIdx As Integer, ByRef colIdx As Integer)
            Dim match = Regex.Match(addr, "^([A-Z]+)([0-9]+)$")
            If Not match.Success Then Throw New Exception(String.Format("Endereço de célula inválido: {0}", addr))
            Dim colStr = match.Groups(1).Value
            Dim rowStr = match.Groups(2).Value
            rowIdx = Integer.Parse(rowStr) - 1
            colIdx = 0
            For Each c In colStr
                colIdx *= 26
                colIdx += (Asc(c) - Asc("A"c) + 1)
            Next
            colIdx -= 1
        End Sub

        Private Function GetVal(dt As DataTable, address As String) As String
            Dim r, c As Integer
            ParseAddress(address, r, c)
            Return GetVal(dt, r, c)
        End Function

        Private Function GetDateVal(dt As DataTable, address As String) As Date?
            Dim r, c As Integer
            ParseAddress(address, r, c)
            Return GetDateFromCell(dt, r, c)
        End Function

        Private Function GetDecimalVal(dt As DataTable, address As String) As Decimal
            Dim r, c As Integer
            ParseAddress(address, r, c)
            Return GetDecimalFromCell(dt, r, c)
        End Function

        Private Function GetVal(dt As DataTable, r As Integer, c As Integer) As String
            If r >= dt.Rows.Count OrElse c >= dt.Columns.Count Then Return ""
            If dt.Rows(r)(c) Is DBNull.Value Then Return ""
            Return dt.Rows(r)(c).ToString().Trim()
        End Function

        Private Function ParseEscolaridade(dt As DataTable) As String
            Dim nivel As String = ""
            Dim status As String = ""
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "D17")) Then nivel = "Ensino Fundamental"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "H17")) Then nivel = "Ensino Médio"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "L17")) Then nivel = "Ensino Superior"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "N17")) Then nivel = "Pós/Especialização"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "D18")) Then status = "Completo"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "G18")) Then status = "Incompleto"
            Dim curso = GetVal(dt, "L18")
            If String.IsNullOrWhiteSpace(nivel) AndAlso String.IsNullOrWhiteSpace(status) Then Return ""
            If String.IsNullOrWhiteSpace(curso) Then
                Return String.Format("{0} {1}", nivel, status).Trim()
            Else
                Return String.Format("{0} {1} - {2}", nivel, status, curso).Trim()
            End If
        End Function

        Private Function ParseEstadoCivil(dt As DataTable) As String
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "D19")) Then Return "Solteiro"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "F19")) Then Return "Casado"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "H19")) Then Return "Viúvo"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "J19")) Then Return "Separado/Divorciado"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "N19")) Then Return "Amasiado"
            Return ""
        End Function

        Private Function ParseVT(dt As DataTable) As String
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "D49")) Then Return "Sim"
            If Not String.IsNullOrWhiteSpace(GetVal(dt, "G49")) Then Return "Não"
            Return ""
        End Function

        Private Function GetDateFromCell(dt As DataTable, r As Integer, c As Integer) As Date?
            If r >= dt.Rows.Count OrElse c >= dt.Columns.Count Then Return Nothing
            Dim obj = dt.Rows(r)(c)
            If obj Is DBNull.Value Then Return Nothing
            If TypeOf obj Is DateTime Then Return CDate(obj)
            If TypeOf obj Is Double Then
                Try
                    Return DateTime.FromOADate(CDbl(obj))
                Catch
                    Return Nothing
                End Try
            End If
            Dim s = obj.ToString()
            If String.IsNullOrWhiteSpace(s) Then Return Nothing
            If Date.TryParse(s, culture, DateTimeStyles.None, Nothing) Then
                Return Date.Parse(s, culture)
            End If
            Return Nothing
        End Function

        Private Function GetDecimalFromCell(dt As DataTable, r As Integer, c As Integer) As Decimal
            If r >= dt.Rows.Count OrElse c >= dt.Columns.Count Then Return 0D
            Dim obj = dt.Rows(r)(c)
            If obj Is DBNull.Value Then Return 0D
            If TypeOf obj Is Decimal Then Return CDec(obj)
            If TypeOf obj Is Double Then Return CDec(obj)
            If TypeOf obj Is Integer OrElse TypeOf obj Is Long Then Return CDec(obj)
            Dim s = obj.ToString().Replace("R$", "").Trim()
            If Decimal.TryParse(s, NumberStyles.Currency, culture, Nothing) Then
                Return Decimal.Parse(s, NumberStyles.Currency, culture)
            End If
            Return 0D
        End Function
    End Module
End Namespace