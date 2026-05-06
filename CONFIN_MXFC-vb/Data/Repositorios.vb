'======================================
' Repositorios.vb
' VERSÃO: 15.0.0 (Migração Web - Otimização Async)
'=======================================
Imports System.Data.SqlClient
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Data
Imports System.Threading.Tasks
Imports System

Namespace Data
    Public Class Periodo
        Public Property Mes As Integer
        Public Property Ano As Integer
        Public ReadOnly Property Label As String
            Get
                Return String.Format("{0:00}/{1}", Mes, Ano)
            End Get
        End Property
    End Class

    ' ==================================================================================
    ' 1. REPOSITÓRIO COLABORADOR
    ' ==================================================================================
    Public Class ColaboradorRepo
        Public Async Function ListarAsync() As Task(Of List(Of Colaborador))
            Dim lst As New List(Of Colaborador)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM Colaborador ORDER BY NomeEmpresa"
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim c As New Colaborador() With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .NomeEmpresa = rd("NomeEmpresa").ToString(),
                            .CNPJ = If(IsDBNull(rd("CNPJ")), "", rd("CNPJ").ToString()),
                            .Endereco = If(IsDBNull(rd("Endereco")), "", rd("Endereco").ToString()),
                            .Cidade = If(IsDBNull(rd("Cidade")), "", rd("Cidade").ToString()),
                            .CEP = If(IsDBNull(rd("CEP")), "", rd("CEP").ToString()),
                            .Responsavel = If(IsDBNull(rd("Responsavel")), "", rd("Responsavel").ToString()),
                            .Telefone = If(IsDBNull(rd("Telefone")), "", rd("Telefone").ToString()),
                            .Email = If(IsDBNull(rd("Email")), "", rd("Email").ToString()),
                            .PixOuDadosBancarios = If(IsDBNull(rd("PixOuDadosBancarios")), "", rd("PixOuDadosBancarios").ToString())
                        }
                        If Not IsDBNull(rd("FuncionarioREG")) Then
                            c.FuncionarioREG = Convert.ToInt32(rd("FuncionarioREG"))
                        End If
                        lst.Add(c)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function SalvarAsync(c As Colaborador) As Task(Of Integer)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If c.Id = 0 Then
                    cmd.CommandText = "INSERT INTO Colaborador(NomeEmpresa,CNPJ,Endereco,Cidade,CEP,Responsavel,Telefone,Email,PixOuDadosBancarios,FuncionarioREG) VALUES(@n,@cnpj,@e,@cid,@cep,@r,@t,@em,@pix,@freg); SELECT CAST(SCOPE_IDENTITY() AS INT);"
                Else
                    cmd.CommandText = "UPDATE Colaborador SET NomeEmpresa=@n,CNPJ=@cnpj,Endereco=@e,Cidade=@cid,CEP=@cep,Responsavel=@r,Telefone=@t,Email=@em,PixOuDadosBancarios=@pix,FuncionarioREG=@freg WHERE Id=@id; SELECT @id;"
                    cmd.Parameters.AddWithValue("@id", c.Id)
                End If
                cmd.Parameters.AddWithValue("@n", c.NomeEmpresa)
                cmd.Parameters.AddWithValue("@cnpj", If(c.CNPJ, ""))
                cmd.Parameters.AddWithValue("@e", If(c.Endereco, ""))
                cmd.Parameters.AddWithValue("@cid", If(c.Cidade, ""))
                cmd.Parameters.AddWithValue("@cep", If(c.CEP, ""))
                cmd.Parameters.AddWithValue("@r", If(c.Responsavel, ""))
                cmd.Parameters.AddWithValue("@t", If(c.Telefone, ""))
                cmd.Parameters.AddWithValue("@em", If(c.Email, ""))
                cmd.Parameters.AddWithValue("@pix", If(c.PixOuDadosBancarios, ""))
                If c.FuncionarioREG.HasValue AndAlso c.FuncionarioREG.Value > 0 Then
                    cmd.Parameters.AddWithValue("@freg", c.FuncionarioREG.Value)
                Else
                    cmd.Parameters.AddWithValue("@freg", DBNull.Value)
                End If
                Dim id = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                If c.Id = 0 Then c.Id = id
                Return c.Id
            End Using
        End Function

        Public Async Function ExcluirAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Colaborador WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 2. REPOSITÓRIO GASTO
    ' ==================================================================================
    Public Class GastoRepo
        Public Async Function ListarAsync(colabId As Integer, mes As Integer?, ano As Integer?) As Task(Of List(Of Gasto))
            Dim lst As New List(Of Gasto)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If mes.HasValue AndAlso ano.HasValue Then
                    Dim dtIni = New Date(ano.Value, mes.Value, 1)
                    Dim dtFim = dtIni.AddMonths(1).AddDays(-1)
                    cmd.CommandText = "SELECT * FROM Gasto WHERE ColaboradorId=@c AND Data BETWEEN @i AND @f ORDER BY Data"
                    cmd.Parameters.AddWithValue("@i", dtIni)
                    cmd.Parameters.AddWithValue("@f", dtFim)
                Else
                    cmd.CommandText = "SELECT * FROM Gasto WHERE ColaboradorId=@c ORDER BY Data"
                End If
                cmd.Parameters.AddWithValue("@c", colabId)
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim g As New Gasto With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .ColaboradorId = Convert.ToInt32(rd("ColaboradorId")),
                            .Data = Convert.ToDateTime(rd("Data")),
                            .NomeDespesa = rd("NomeDespesa").ToString(),
                            .TipoDespesa = rd("TipoDespesa").ToString(),
                            .Valor = CDec(Convert.ToDouble(rd("Valor")))
                        }
                        Dim idxAnexo = rd.GetOrdinal("Anexo")
                        If Not rd.IsDBNull(idxAnexo) Then
                            g.Anexo = DirectCast(rd(idxAnexo), Byte())
                        End If
                        lst.Add(g)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function SalvarAsync(g As Gasto) As Task(Of Integer)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If g.Id = 0 Then
                    cmd.CommandText = "INSERT INTO Gasto(ColaboradorId,Data,NomeDespesa,TipoDespesa,Valor,Anexo) VALUES(@c,@d,@n,@t,@v,@anexo); SELECT CAST(SCOPE_IDENTITY() AS INT);"
                Else
                    cmd.CommandText = "UPDATE Gasto SET ColaboradorId=@c,Data=@d,NomeDespesa=@n,TipoDespesa=@t,Valor=@v,Anexo=@anexo WHERE Id=@id; SELECT @id;"
                    cmd.Parameters.AddWithValue("@id", g.Id)
                End If
                cmd.Parameters.AddWithValue("@c", g.ColaboradorId)
                cmd.Parameters.AddWithValue("@d", g.Data)
                cmd.Parameters.AddWithValue("@n", g.NomeDespesa)
                cmd.Parameters.AddWithValue("@t", g.TipoDespesa)
                cmd.Parameters.AddWithValue("@v", g.Valor)
                Dim paramAnexo As New SqlParameter("@anexo", SqlDbType.VarBinary)
                If g.Anexo IsNot Nothing AndAlso g.Anexo.Length > 0 Then
                    paramAnexo.Value = g.Anexo
                Else
                    paramAnexo.Value = DBNull.Value
                End If
                cmd.Parameters.Add(paramAnexo)
                Dim id = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                Return If(g.Id = 0, id, g.Id)
            End Using
        End Function

        Public Async Function ExcluirAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Gasto WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function SomarDoPeriodoAsync(colabId As Integer, mes As Integer, ano As Integer) As Task(Of Decimal)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                Dim dtBase = New Date(ano, mes, 1)
                Dim dtInicioGastos = dtBase.AddMonths(-1)
                Dim dtFimGastos = dtInicioGastos.AddMonths(1).AddDays(-1)
                cmd.CommandText = "SELECT ISNULL(SUM(Valor),0) FROM Gasto WHERE ColaboradorId=@c AND Data BETWEEN @i AND @f"
                cmd.Parameters.AddWithValue("@c", colabId)
                cmd.Parameters.AddWithValue("@i", dtInicioGastos)
                cmd.Parameters.AddWithValue("@f", dtFimGastos)
                Return CDec(Convert.ToDouble(Await cmd.ExecuteScalarAsync()))
            End Using
        End Function

        Public Async Function ListarNomesDespesaAsync() As Task(Of List(Of String))
            Dim lst As New List(Of String)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT DISTINCT NomeDespesa FROM Gasto WHERE NomeDespesa IS NOT NULL AND NomeDespesa <> '' ORDER BY NomeDespesa"
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        lst.Add(rd.GetString(0))
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function ListarPeriodosDisponiveisAsync() As Task(Of List(Of Periodo))
            Dim lst As New List(Of Periodo)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT DISTINCT MONTH(Data) AS Mes, YEAR(Data) AS Ano FROM Gasto UNION SELECT DISTINCT Mes, Ano FROM NotaFiscal ORDER BY Ano DESC, Mes DESC"
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        lst.Add(New Periodo With {.Mes = rd.GetInt32(0), .Ano = rd.GetInt32(1)})
                    End While
                End Using
            End Using
            Return lst
        End Function
    End Class

    ' ==================================================================================
    ' 3. REPOSITÓRIO NOTA FISCAL
    ' ==================================================================================
    Public Class NotaFiscalRepo
        Public Async Function ObterOuCriarAsync(colabId As Integer, mes As Integer, ano As Integer) As Task(Of FinanceiroPJ.Models.NotaFiscal)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM NotaFiscal WHERE ColaboradorId=@c AND Mes=@m AND Ano=@a"
                cmd.Parameters.AddWithValue("@c", colabId)
                cmd.Parameters.AddWithValue("@m", mes)
                cmd.Parameters.AddWithValue("@a", ano)
                Using rd = Await cmd.ExecuteReaderAsync()
                    If Await rd.ReadAsync() Then
                        Dim adicionalVal As Decimal = 0
                        If Not rd.IsDBNull(rd.GetOrdinal("Adicional")) Then adicionalVal = CDec(Convert.ToDouble(rd("Adicional")))
                        Dim ccVal As String = ""
                        If Not rd.IsDBNull(rd.GetOrdinal("CentroCusto")) Then ccVal = rd("CentroCusto").ToString()
                        Return New FinanceiroPJ.Models.NotaFiscal With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .ColaboradorId = Convert.ToInt32(rd("ColaboradorId")),
                            .Mes = Convert.ToInt32(rd("Mes")),
                            .Ano = Convert.ToInt32(rd("Ano")),
                            .ValorFixo = CDec(Convert.ToDouble(rd("ValorFixo"))),
                            .FG = CDec(Convert.ToDouble(rd("FG"))),
                            .Adicional = adicionalVal,
                            .CentroCusto = ccVal,
                            .Despesas = CDec(Convert.ToDouble(rd("Despesas"))),
                            .Contador = CDec(Convert.ToDouble(rd("Contador"))),
                            .INSSProlabore = CDec(Convert.ToDouble(rd("INSSProlabore"))),
                            .ImpostoNF = CDec(Convert.ToDouble(rd("ImpostoNF"))),
                            .ValorNF = CDec(Convert.ToDouble(rd("ValorNF"))),
                            .Prolabore = CDec(Convert.ToDouble(rd("Prolabore"))),
                            .Observacao = If(rd.IsDBNull(rd.GetOrdinal("Observacao")), "", rd("Observacao").ToString())
                        }
                    End If
                End Using
            End Using
            Return New FinanceiroPJ.Models.NotaFiscal With {.ColaboradorId = colabId, .Mes = mes, .Ano = ano}
        End Function

        Public Async Function SalvarAsync(nf As FinanceiroPJ.Models.NotaFiscal) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If nf.Id = 0 Then
                    cmd.CommandText = "INSERT INTO NotaFiscal(ColaboradorId,Mes,Ano,ValorFixo,FG,Adicional,CentroCusto,Despesas,Contador,INSSProlabore,ImpostoNF,ValorNF,Prolabore,Observacao) VALUES(@c,@m,@a,@vf,@fg,@add,@cc,@d,@cont,@inss,@imp,@v,@p,@o)"
                Else
                    cmd.CommandText = "UPDATE NotaFiscal SET ValorFixo=@vf,FG=@fg,Adicional=@add,CentroCusto=@cc,Despesas=@d,Contador=@cont,INSSProlabore=@inss,ImpostoNF=@imp,ValorNF=@v,Prolabore=@p,Observacao=@o WHERE Id=@id"
                    cmd.Parameters.AddWithValue("@id", nf.Id)
                End If
                cmd.Parameters.AddWithValue("@c", nf.ColaboradorId)
                cmd.Parameters.AddWithValue("@m", nf.Mes)
                cmd.Parameters.AddWithValue("@a", nf.Ano)
                cmd.Parameters.AddWithValue("@vf", nf.ValorFixo)
                cmd.Parameters.AddWithValue("@fg", nf.FG)
                cmd.Parameters.AddWithValue("@add", nf.Adicional)
                cmd.Parameters.AddWithValue("@cc", If(nf.CentroCusto, DBNull.Value))
                cmd.Parameters.AddWithValue("@d", nf.Despesas)
                cmd.Parameters.AddWithValue("@cont", nf.Contador)
                cmd.Parameters.AddWithValue("@inss", nf.INSSProlabore)
                cmd.Parameters.AddWithValue("@imp", nf.ImpostoNF)
                cmd.Parameters.AddWithValue("@v", nf.ValorNF)
                cmd.Parameters.AddWithValue("@p", nf.Prolabore)
                cmd.Parameters.AddWithValue("@o", If(nf.Observacao, DBNull.Value))
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 4. REPOSITÓRIO CLIENTE
    ' ==================================================================================
    Public Class ClienteRepo
        Public Async Function ListarAsync() As Task(Of List(Of Cliente))
            Dim lst As New List(Of Cliente)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM Cliente ORDER BY RazaoSocial"
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim c As New Cliente With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .RazaoSocial = rd("RazaoSocial").ToString(),
                            .CNPJ = If(rd.IsDBNull(rd.GetOrdinal("CNPJ")), "", rd("CNPJ").ToString()),
                            .InscricaoEstadual = If(rd.IsDBNull(rd.GetOrdinal("InscricaoEstadual")), "", rd("InscricaoEstadual").ToString()),
                            .Endereco = If(rd.IsDBNull(rd.GetOrdinal("Endereco")), "", rd("Endereco").ToString()),
                            .CEP = If(rd.IsDBNull(rd.GetOrdinal("CEP")), "", rd("CEP").ToString()),
                            .Telefone = If(rd.IsDBNull(rd.GetOrdinal("Telefone")), "", rd("Telefone").ToString()),
                            .Responsavel = If(rd.IsDBNull(rd.GetOrdinal("Responsavel")), "", rd("Responsavel").ToString())
                        }
                        lst.Add(c)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function SalvarAsync(c As Cliente) As Task(Of Integer)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If c.Id = 0 Then
                    cmd.CommandText = "INSERT INTO Cliente(RazaoSocial,CNPJ,InscricaoEstadual,Endereco,CEP,Telefone,Responsavel) VALUES(@rs,@cnpj,@ie,@end,@cep,@tel,@resp); SELECT CAST(SCOPE_IDENTITY() AS INT);"
                Else
                    cmd.CommandText = "UPDATE Cliente SET RazaoSocial=@rs,CNPJ=@cnpj,InscricaoEstadual=@ie,Endereco=@end,CEP=@cep,Telefone=@tel,Responsavel=@resp WHERE Id=@id; SELECT @id;"
                    cmd.Parameters.AddWithValue("@id", c.Id)
                End If
                cmd.Parameters.AddWithValue("@rs", c.RazaoSocial)
                cmd.Parameters.AddWithValue("@cnpj", If(c.CNPJ, ""))
                cmd.Parameters.AddWithValue("@ie", If(c.InscricaoEstadual, ""))
                cmd.Parameters.AddWithValue("@end", If(c.Endereco, ""))
                cmd.Parameters.AddWithValue("@cep", If(c.CEP, ""))
                cmd.Parameters.AddWithValue("@tel", If(c.Telefone, ""))
                cmd.Parameters.AddWithValue("@resp", If(c.Responsavel, ""))
                Dim id = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                If c.Id = 0 Then c.Id = id
                Return c.Id
            End Using
        End Function

        Public Async Function ExcluirAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Cliente WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 5. REPOSITÓRIO CONTRATO
    ' ==================================================================================
    Public Class ContratoRepo
        Public Async Function ListarAsync(clienteId As Integer) As Task(Of List(Of Contrato))
            Dim lst As New List(Of Contrato)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT c.*, ISNULL((SELECT SUM(Valor) FROM ContratoAditivo WHERE ContratoId = c.Id), 0) AS ValorAditivos, (SELECT MAX(DataFinal) FROM ContratoAditivo WHERE ContratoId = c.Id) AS DataFinalAditivo FROM Contrato c WHERE ClienteId=@cid ORDER BY DataInicial DESC"
                cmd.Parameters.AddWithValue("@cid", clienteId)
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim c As New Contrato With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .ClienteId = Convert.ToInt32(rd("ClienteId")),
                            .NumeroContrato = rd("NumeroContrato").ToString(),
                            .Valor = CDec(Convert.ToDouble(rd("Valor"))),
                            .DataInicial = If(rd.IsDBNull(rd.GetOrdinal("DataInicial")), Date.MinValue, Convert.ToDateTime(rd("DataInicial"))),
                            .CentroCusto = If(rd.IsDBNull(rd.GetOrdinal("CentroCusto")), "", rd("CentroCusto").ToString()),
                            .ValorAditivos = CDec(Convert.ToDouble(rd("ValorAditivos")))
                        }
                        Dim dtBase = If(rd.IsDBNull(rd.GetOrdinal("DataFinal")), Date.MinValue, Convert.ToDateTime(rd("DataFinal")))
                        Dim dtAditivo = If(rd.IsDBNull(rd.GetOrdinal("DataFinalAditivo")), Date.MinValue, Convert.ToDateTime(rd("DataFinalAditivo")))
                        If dtAditivo > dtBase Then
                            c.DataFinal = dtAditivo
                        Else
                            c.DataFinal = dtBase
                        End If
                        If c.DataInicial <> Date.MinValue AndAlso c.DataFinal <> Date.MinValue Then
                            c.Prazo = CInt((c.DataFinal - c.DataInicial).TotalDays)
                        Else
                            c.Prazo = If(rd.IsDBNull(rd.GetOrdinal("Prazo")), 0, Convert.ToInt32(rd("Prazo")))
                        End If
                        lst.Add(c)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function SalvarAsync(c As Contrato) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                Dim contratoId As Integer
                If c.Id = 0 Then
                    cmd.CommandText = "INSERT INTO Contrato(ClienteId,NumeroContrato,Valor,DataInicial,DataFinal,Prazo,CentroCusto) VALUES(@cid,@num,@val,@di,@df,@pz,@cc); SELECT CAST(SCOPE_IDENTITY() AS INT);"
                Else
                    cmd.CommandText = "UPDATE Contrato SET ClienteId=@cid,NumeroContrato=@num,Valor=@val,DataInicial=@di,DataFinal=@df,Prazo=@pz,CentroCusto=@cc WHERE Id=@id; SELECT @id;"
                    cmd.Parameters.AddWithValue("@id", c.Id)
                End If
                cmd.Parameters.AddWithValue("@cid", c.ClienteId)
                cmd.Parameters.AddWithValue("@num", c.NumeroContrato)
                cmd.Parameters.AddWithValue("@val", c.Valor)
                cmd.Parameters.AddWithValue("@di", c.DataInicial)
                cmd.Parameters.AddWithValue("@df", c.DataFinal)
                cmd.Parameters.AddWithValue("@pz", c.Prazo)
                cmd.Parameters.AddWithValue("@cc", If(c.CentroCusto, DBNull.Value))
                contratoId = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                c.Id = contratoId
                If c.Anexos IsNot Nothing AndAlso c.Anexos.Count > 0 Then
                    For Each anexo In c.Anexos
                        If anexo.Id = 0 Then
                            Using cmdAnexo = cn.CreateCommand()
                                cmdAnexo.CommandText = "INSERT INTO ContratoAnexo(ContratoId, NomeArquivo, Conteudo) VALUES(@cid, @nome, @conteudo)"
                                cmdAnexo.Parameters.AddWithValue("@cid", contratoId)
                                cmdAnexo.Parameters.AddWithValue("@nome", anexo.NomeArquivo)
                                cmdAnexo.Parameters.AddWithValue("@conteudo", anexo.Conteudo)
                                Await cmdAnexo.ExecuteNonQueryAsync()
                            End Using
                        End If
                    Next
                End If
            End Using
        End Function

        Public Async Function ListarAnexosAsync(contratoId As Integer) As Task(Of List(Of ContratoAnexo))
            Dim lst As New List(Of ContratoAnexo)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT Id, NomeArquivo, Conteudo FROM ContratoAnexo WHERE ContratoId=@cid"
                cmd.Parameters.AddWithValue("@cid", contratoId)
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        lst.Add(New ContratoAnexo With {
                            .Id = rd.GetInt32(0),
                            .ContratoId = contratoId,
                            .NomeArquivo = rd.GetString(1),
                            .Conteudo = DirectCast(rd(2), Byte())
                        })
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function ExcluirAnexoAsync(anexoId As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM ContratoAnexo WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", anexoId)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function ExcluirAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Contrato WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 6. REPOSITÓRIO CONTRATO ADITIVO
    ' ==================================================================================
    Public Class ContratoAditivoRepo
        Public Async Function ListarAsync(contratoId As Integer) As Task(Of List(Of ContratoAditivo))
            Dim lst As New List(Of ContratoAditivo)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT Id, ContratoId, Numero, Data, DataFinal, Valor, Descricao, CASE WHEN Anexo IS NOT NULL THEN 1 ELSE 0 END as TemAnexo, Anexo FROM ContratoAditivo WHERE ContratoId=@cid ORDER BY Data DESC"
                cmd.Parameters.AddWithValue("@cid", contratoId)
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim a As New ContratoAditivo With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .ContratoId = Convert.ToInt32(rd("ContratoId")),
                            .Numero = If(rd.IsDBNull(rd.GetOrdinal("Numero")), "", rd("Numero").ToString()),
                            .Data = Convert.ToDateTime(rd("Data")),
                            .DataFinal = If(rd.IsDBNull(rd.GetOrdinal("DataFinal")), Date.MinValue, Convert.ToDateTime(rd("DataFinal"))),
                            .Valor = CDec(Convert.ToDouble(rd("Valor"))),
                            .Descricao = If(rd.IsDBNull(rd.GetOrdinal("Descricao")), "", rd("Descricao").ToString()),
                            .TemAnexo = (Convert.ToInt32(rd("TemAnexo")) = 1)
                        }
                        If a.TemAnexo Then
                            a.Anexo = DirectCast(rd("Anexo"), Byte())
                        End If
                        lst.Add(a)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function SalvarAsync(a As ContratoAditivo) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                If a.Id = 0 Then
                    cmd.CommandText = "INSERT INTO ContratoAditivo(ContratoId,Numero,Data,DataFinal,Valor,Descricao,Anexo) VALUES(@cid,@num,@dt,@df,@val,@desc,@anexo)"
                Else
                    cmd.CommandText = "UPDATE ContratoAditivo SET ContratoId=@cid,Numero=@num,Data=@dt,DataFinal=@df,Valor=@val,Descricao=@desc,Anexo=@anexo WHERE Id=@id"
                    cmd.Parameters.AddWithValue("@id", a.Id)
                End If
                cmd.Parameters.AddWithValue("@cid", a.ContratoId)
                cmd.Parameters.AddWithValue("@num", a.Numero)
                cmd.Parameters.AddWithValue("@dt", a.Data)
                cmd.Parameters.AddWithValue("@df", If(a.DataFinal = Date.MinValue, DBNull.Value, a.DataFinal))
                cmd.Parameters.AddWithValue("@val", a.Valor)
                cmd.Parameters.AddWithValue("@desc", If(a.Descricao, DBNull.Value))
                Dim pAnexo As New SqlParameter("@anexo", SqlDbType.VarBinary)
                If a.Anexo IsNot Nothing AndAlso a.Anexo.Length > 0 Then
                    pAnexo.Value = a.Anexo
                Else
                    pAnexo.Value = DBNull.Value
                End If
                cmd.Parameters.Add(pAnexo)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function ExcluirAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM ContratoAditivo WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 7. REPOSITÓRIO FUNCIONÁRIO
    ' ==================================================================================
    Public Class FuncionarioRepo
        Private Shared Function GetStr(rd As SqlDataReader, colName As String) As String
            Try
                Dim idx = rd.GetOrdinal(colName)
                If rd.IsDBNull(idx) Then Return ""
                Return rd.GetString(idx)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetBytes(rd As SqlDataReader, colName As String) As Byte()
            Try
                Dim idx = rd.GetOrdinal(colName)
                If rd.IsDBNull(idx) Then Return Nothing
                Return DirectCast(rd(idx), Byte())
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function GetDate(rd As SqlDataReader, colName As String) As Date?
            Try
                Dim idx = rd.GetOrdinal(colName)
                If rd.IsDBNull(idx) Then Return Nothing
                Dim val = rd.GetValue(idx).ToString()
                If String.IsNullOrWhiteSpace(val) Then Return Nothing
                Dim dt As Date
                If Date.TryParseExact(val, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then Return dt
                If Date.TryParse(val, dt) Then Return dt
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function GetDec(rd As SqlDataReader, colName As String) As Decimal
            Try
                Dim idx = rd.GetOrdinal(colName)
                If rd.IsDBNull(idx) Then Return 0D
                Return CDec(Convert.ToDouble(rd(idx)))
            Catch
                Return 0D
            End Try
        End Function

        Private Shared Function GetBool(rd As SqlDataReader, colName As String) As Boolean
            Try
                Dim idx = rd.GetOrdinal(colName)
                If rd.IsDBNull(idx) Then Return True
                Dim val = rd.GetValue(idx)
                If TypeOf val Is Boolean Then Return DirectCast(val, Boolean)
                If val.ToString() = "1" Then Return True
                Return False
            Catch
                Return True
            End Try
        End Function

        Public Async Function AtualizarCaminhoDeclaracaoAsync(reg As Integer, dadosDeclaracao As Byte()) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE Funcionario SET Declaracao = @decl WHERE REG = @reg"
                Dim p As New SqlParameter("@decl", SqlDbType.VarBinary)
                If dadosDeclaracao IsNot Nothing AndAlso dadosDeclaracao.Length > 0 Then
                    p.Value = dadosDeclaracao
                Else
                    p.Value = DBNull.Value
                End If
                cmd.Parameters.Add(p)
                cmd.Parameters.AddWithValue("@reg", reg)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Private Shared Function ReadFuncionario(rd As SqlDataReader) As Funcionario
            Return New Funcionario With {
                .REG = Convert.ToInt32(rd("REG")),
                .NomeCompleto = GetStr(rd, "NomeCompleto"),
                .Foto = GetBytes(rd, "Foto"),
                .Declaracao = GetBytes(rd, "Declaracao"),
                .Situacao = GetStr(rd, "Situacao"),
                .DataNascimento = GetDate(rd, "DataNascimento"),
                .Naturalidade = GetStr(rd, "Naturalidade"),
                .Nacionalidade = GetStr(rd, "Nacionalidade"),
                .EstadoCivil = GetStr(rd, "EstadoCivil"),
                .NomeConjuge = GetStr(rd, "NomeConjuge"),
                .NomePai = GetStr(rd, "NomePai"),
                .NomeMae = GetStr(rd, "NomeMae"),
                .Endereco = GetStr(rd, "Endereco"),
                .EnderecoNumero = GetStr(rd, "EnderecoNumero"),
                .EnderecoComplemento = GetStr(rd, "EnderecoComplemento"),
                .EnderecoBairro = GetStr(rd, "EnderecoBairro"),
                .CEP = GetStr(rd, "CEP"),
                .Cidade = GetStr(rd, "Cidade"),
                .Estado = GetStr(rd, "Estado"),
                .Telefone = GetStr(rd, "Telefone"),
                .Celular = GetStr(rd, "Celular"),
                .Email = GetStr(rd, "Email"),
                .Email2 = GetStr(rd, "Email2"),
                .Formacao = GetStr(rd, "Formacao"),
                .Cargo = GetStr(rd, "Cargo"),
                .DataAdmissao = GetDate(rd, "DataAdmissao"),
                .RegimeContratacao = GetStr(rd, "RegimeContratacao"),
                .Salario = GetDec(rd, "Salario"),
                .PrazoExperiencia = GetStr(rd, "PrazoExperiencia"),
                .HorarioTrabalho = GetStr(rd, "HorarioTrabalho"),
                .DescontaVT = GetStr(rd, "DescontaVT"),
                .ObservacoesRH = GetStr(rd, "ObservacoesRH"),
                .RG = GetStr(rd, "RG"),
                .RGDataExpedicao = GetDate(rd, "RGDataExpedicao"),
                .RGOrgaoExp = GetStr(rd, "RGOrgaoExp"),
                .RGUF = GetStr(rd, "RGUF"),
                .CPF = GetStr(rd, "CPF"),
                .PIS = GetStr(rd, "PIS"),
                .CTPS = GetStr(rd, "CTPS"),
                .CTPSSerie = GetStr(rd, "CTPSSerie"),
                .CTPSUF = GetStr(rd, "CTPSUF"),
                .CTPSDataExp = GetDate(rd, "CTPSDataExp"),
                .TituloEleitor = GetStr(rd, "TituloEleitor"),
                .TituloZona = GetStr(rd, "TituloZona"),
                .TituloSecao = GetStr(rd, "TituloSecao"),
                .TituloMunicipioUF = GetStr(rd, "TituloMunicipioUF"),
                .TituloDataEmissao = GetDate(rd, "TituloDataEmissao"),
                .Reservista = GetStr(rd, "Reservista"),
                .CNH = GetStr(rd, "CNH"),
                .CNHCategoria = GetStr(rd, "CNHCategoria"),
                .CNHValidade = GetDate(rd, "CNHValidade"),
                .CNHData1Hab = GetDate(rd, "CNHData1Hab"),
                .CNHDataEmissao = GetDate(rd, "CNHDataEmissao"),
                .ConjugeDataNasc = GetDate(rd, "ConjugeDataNasc"),
                .ConjugeRG = GetStr(rd, "ConjugeRG"),
                .ConjugeRGUF = GetStr(rd, "ConjugeRGUF"),
                .ConjugeCPF = GetStr(rd, "ConjugeCPF"),
                .Banco = GetStr(rd, "Banco"),
                .Agencia = GetStr(rd, "Agencia"),
                .ContaCorrente = GetStr(rd, "ContaCorrente"),
                .ContaModalidade = GetStr(rd, "ContaModalidade"),
                .PIX = GetStr(rd, "PIX"),
                .Banco2 = GetStr(rd, "Banco2"),
                .Agencia2 = GetStr(rd, "Agencia2"),
                .ContaCorrente2 = GetStr(rd, "ContaCorrente2"),
                .ContaModalidade2 = GetStr(rd, "ContaModalidade2"),
                .PIX2 = GetStr(rd, "PIX2"),
                .Dependentes = GetStr(rd, "Dependentes"),
                .EmergenciaNome = GetStr(rd, "EmergenciaNome"),
                .EmergenciaFone = GetStr(rd, "EmergenciaFone"),
                .EmergenciaParentesco = GetStr(rd, "EmergenciaParentesco"),
                .Ativo = GetBool(rd, "Ativo")
            }
        End Function

        Public Async Function ListarAsync() As Task(Of List(Of Funcionario))
            Dim lst As New List(Of Funcionario)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM Funcionario ORDER BY NomeCompleto"
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        lst.Add(ReadFuncionario(rd))
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function ObterAsync(reg As Integer) As Task(Of Funcionario)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM Funcionario WHERE REG=@reg"
                cmd.Parameters.AddWithValue("@reg", reg)
                Using rd = Await cmd.ExecuteReaderAsync()
                    If Await rd.ReadAsync() Then
                        Return ReadFuncionario(rd)
                    End If
                End Using
            End Using
            Return Nothing
        End Function

        Public Async Function SalvarAsync(f As Funcionario, Optional reutilizarReg As Boolean = False) As Task(Of Integer)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                Dim usarRegEspecifico As Boolean = False
                If f.REG = 0 Then
                    If reutilizarReg Then
                        cmd.CommandText = "SELECT REG FROM Funcionario ORDER BY REG ASC"
                        Dim regsExistentes As New List(Of Integer)
                        Using rd = Await cmd.ExecuteReaderAsync()
                            While Await rd.ReadAsync()
                                regsExistentes.Add(rd.GetInt32(0))
                            End While
                        End Using
                        Dim candidato As Integer = 1
                        For Each r In regsExistentes
                            If r = candidato Then
                                candidato += 1
                            Else
                                Exit For
                            End If
                        Next
                        f.REG = candidato
                        usarRegEspecifico = True
                        cmd.Parameters.Clear()
                    End If
                    If usarRegEspecifico Then
                        cmd.CommandText = "SET IDENTITY_INSERT Funcionario ON;" &
                                          "INSERT INTO Funcionario (REG, NomeCompleto, Foto, Situacao, DataNascimento, Naturalidade, Nacionalidade, EstadoCivil, NomeConjuge, NomePai, NomeMae, Endereco, EnderecoNumero, EnderecoComplemento, EnderecoBairro, CEP, Cidade, Estado, Telefone, Celular, Email, Email2, Formacao, Cargo, DataAdmissao, RegimeContratacao, Salario, PrazoExperiencia, HorarioTrabalho, DescontaVT, ObservacoesRH, RG, RGDataExpedicao, RGOrgaoExp, RGUF, CPF, PIS, CTPS, CTPSSerie, CTPSUF, CTPSDataExp, TituloEleitor, TituloZona, TituloSecao, TituloMunicipioUF, TituloDataEmissao, Reservista, CNH, CNHCategoria, CNHValidade, CNHData1Hab, CNHDataEmissao, ConjugeDataNasc, ConjugeRG, ConjugeRGUF, ConjugeCPF, Banco, Agencia, ContaCorrente, ContaModalidade, PIX, Banco2, Agencia2, ContaCorrente2, ContaModalidade2, PIX2, Dependentes, EmergenciaNome, EmergenciaFone, EmergenciaParentesco, Ativo) " &
                                          "VALUES(@REG, @NomeCompleto, @Foto, @Situacao, @DataNascimento, @Naturalidade, @Nacionalidade, @EstadoCivil, @NomeConjuge, @NomePai, @NomeMae, @Endereco, @EnderecoNumero, @EnderecoComplemento, @EnderecoBairro, @CEP, @Cidade, @Estado, @Telefone, @Celular, @Email, @Email2, @Formacao, @Cargo, @DataAdmissao, @RegimeContratacao, @Salario, @PrazoExperiencia, @HorarioTrabalho, @DescontaVT, @ObservacoesRH, @RG, @RGDataExpedicao, @RGOrgaoExp, @RGUF, @CPF, @PIS, @CTPS, @CTPSSerie, @CTPSUF, @CTPSDataExp, @TituloEleitor, @TituloZona, @TituloSecao, @TituloMunicipioUF, @TituloDataEmissao, @Reservista, @CNH, @CNHCategoria, @CNHValidade, @CNHData1Hab, @CNHDataEmissao, @ConjugeDataNasc, @ConjugeRG, @ConjugeRGUF, @ConjugeCPF, @Banco, @Agencia, @ContaCorrente, @ContaModalidade, @PIX, @Banco2, @Agencia2, @ContaCorrente2, @ContaModalidade2, @PIX2, @Dependentes, @EmergenciaNome, @EmergenciaFone, @EmergenciaParentesco, @Ativo);" &
                                          "SET IDENTITY_INSERT Funcionario OFF;"
                        cmd.Parameters.AddWithValue("@REG", f.REG)
                    Else
                        cmd.CommandText = "INSERT INTO Funcionario (NomeCompleto, Foto, Situacao, DataNascimento, Naturalidade, Nacionalidade, EstadoCivil, NomeConjuge, NomePai, NomeMae, Endereco, EnderecoNumero, EnderecoComplemento, EnderecoBairro, CEP, Cidade, Estado, Telefone, Celular, Email, Email2, Formacao, Cargo, DataAdmissao, RegimeContratacao, Salario, PrazoExperiencia, HorarioTrabalho, DescontaVT, ObservacoesRH, RG, RGDataExpedicao, RGOrgaoExp, RGUF, CPF, PIS, CTPS, CTPSSerie, CTPSUF, CTPSDataExp, TituloEleitor, TituloZona, TituloSecao, TituloMunicipioUF, TituloDataEmissao, Reservista, CNH, CNHCategoria, CNHValidade, CNHData1Hab, CNHDataEmissao, ConjugeDataNasc, ConjugeRG, ConjugeRGUF, ConjugeCPF, Banco, Agencia, ContaCorrente, ContaModalidade, PIX, Banco2, Agencia2, ContaCorrente2, ContaModalidade2, PIX2, Dependentes, EmergenciaNome, EmergenciaFone, EmergenciaParentesco, Ativo) " &
                                          "VALUES(@NomeCompleto, @Foto, @Situacao, @DataNascimento, @Naturalidade, @Nacionalidade, @EstadoCivil, @NomeConjuge, @NomePai, @NomeMae, @Endereco, @EnderecoNumero, @EnderecoComplemento, @EnderecoBairro, @CEP, @Cidade, @Estado, @Telefone, @Celular, @Email, @Email2, @Formacao, @Cargo, @DataAdmissao, @RegimeContratacao, @Salario, @PrazoExperiencia, @HorarioTrabalho, @DescontaVT, @ObservacoesRH, @RG, @RGDataExpedicao, @RGOrgaoExp, @RGUF, @CPF, @PIS, @CTPS, @CTPSSerie, @CTPSUF, @CTPSDataExp, @TituloEleitor, @TituloZona, @TituloSecao, @TituloMunicipioUF, @TituloDataEmissao, @Reservista, @CNH, @CNHCategoria, @CNHValidade, @CNHData1Hab, @CNHDataEmissao, @ConjugeDataNasc, @ConjugeRG, @ConjugeRGUF, @ConjugeCPF, @Banco, @Agencia, @ContaCorrente, @ContaModalidade, @PIX, @Banco2, @Agencia2, @ContaCorrente2, @ContaModalidade2, @PIX2, @Dependentes, @EmergenciaNome, @EmergenciaFone, @EmergenciaParentesco, @Ativo);" &
                                          "SELECT CAST(SCOPE_IDENTITY() AS INT);"
                    End If
                Else
                    cmd.CommandText = "UPDATE Funcionario SET NomeCompleto=@NomeCompleto, Foto=@Foto, Situacao=@Situacao, DataNascimento=@DataNascimento, Naturalidade=@Naturalidade, Nacionalidade=@Nacionalidade, EstadoCivil=@EstadoCivil, NomeConjuge=@NomeConjuge, NomePai=@NomePai, NomeMae=@NomeMae, Endereco=@Endereco, EnderecoNumero=@EnderecoNumero, EnderecoComplemento=@EnderecoComplemento, EnderecoBairro=@EnderecoBairro, CEP=@CEP, Cidade=@Cidade, Estado=@Estado, Telefone=@Telefone, Celular=@Celular, Email=@Email, Email2=@Email2, Formacao=@Formacao, Cargo=@Cargo, DataAdmissao=@DataAdmissao, RegimeContratacao=@RegimeContratacao, Salario=@Salario, PrazoExperiencia=@PrazoExperiencia, HorarioTrabalho=@HorarioTrabalho, DescontaVT=@DescontaVT, ObservacoesRH=@ObservacoesRH, RG=@RG, RGDataExpedicao=@RGDataExpedicao, RGOrgaoExp=@RGOrgaoExp, RGUF=@RGUF, CPF=@CPF, PIS=@PIS, CTPS=@CTPS, CTPSSerie=@CTPSSerie, CTPSUF=@CTPSUF, CTPSDataExp=@CTPSDataExp, TituloEleitor=@TituloEleitor, TituloZona=@TituloZona, TituloSecao=@TituloSecao, TituloMunicipioUF=@TituloMunicipioUF, TituloDataEmissao=@TituloDataEmissao, Reservista=@Reservista, CNH=@CNH, CNHCategoria=@CNHCategoria, CNHValidade=@CNHValidade, CNHData1Hab=@CNHData1Hab, CNHDataEmissao=@CNHDataEmissao, ConjugeDataNasc=@ConjugeDataNasc, ConjugeRG=@ConjugeRG, ConjugeRGUF=@ConjugeRGUF, ConjugeCPF=@ConjugeCPF, Banco=@Banco, Agencia=@Agencia, ContaCorrente=@ContaCorrente, ContaModalidade=@ContaModalidade, PIX=@PIX, Banco2=@Banco2, Agencia2=@Agencia2, ContaCorrente2=@ContaCorrente2, ContaModalidade2=@ContaModalidade2, PIX2=@PIX2, Dependentes=@Dependentes, EmergenciaNome=@EmergenciaNome, EmergenciaFone=@EmergenciaFone, EmergenciaParentesco=@EmergenciaParentesco, Ativo=@Ativo " &
                                      "WHERE REG=@REG; SELECT @REG;"
                    cmd.Parameters.AddWithValue("@REG", f.REG)
                End If
                cmd.Parameters.AddWithValue("@NomeCompleto", f.NomeCompleto)
                Dim pFoto As New SqlParameter("@Foto", SqlDbType.VarBinary)
                If f.Foto IsNot Nothing AndAlso f.Foto.Length > 0 Then
                    pFoto.Value = f.Foto
                Else
                    pFoto.Value = DBNull.Value
                End If
                cmd.Parameters.Add(pFoto)
                cmd.Parameters.AddWithValue("@Situacao", If(String.IsNullOrEmpty(f.Situacao), "Ativo", f.Situacao))
                cmd.Parameters.AddWithValue("@DataNascimento", If(f.DataNascimento.HasValue, f.DataNascimento.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@Naturalidade", If(f.Naturalidade, DBNull.Value))
                cmd.Parameters.AddWithValue("@Nacionalidade", If(f.Nacionalidade, DBNull.Value))
                cmd.Parameters.AddWithValue("@EstadoCivil", If(f.EstadoCivil, DBNull.Value))
                cmd.Parameters.AddWithValue("@NomeConjuge", If(f.NomeConjuge, DBNull.Value))
                cmd.Parameters.AddWithValue("@NomePai", If(f.NomePai, DBNull.Value))
                cmd.Parameters.AddWithValue("@NomeMae", If(f.NomeMae, DBNull.Value))
                cmd.Parameters.AddWithValue("@Endereco", If(f.Endereco, DBNull.Value))
                cmd.Parameters.AddWithValue("@EnderecoNumero", If(f.EnderecoNumero, DBNull.Value))
                cmd.Parameters.AddWithValue("@EnderecoComplemento", If(f.EnderecoComplemento, DBNull.Value))
                cmd.Parameters.AddWithValue("@EnderecoBairro", If(f.EnderecoBairro, DBNull.Value))
                cmd.Parameters.AddWithValue("@CEP", If(f.CEP, DBNull.Value))
                cmd.Parameters.AddWithValue("@Cidade", If(f.Cidade, DBNull.Value))
                cmd.Parameters.AddWithValue("@Estado", If(f.Estado, DBNull.Value))
                cmd.Parameters.AddWithValue("@Telefone", If(f.Telefone, DBNull.Value))
                cmd.Parameters.AddWithValue("@Celular", If(f.Celular, DBNull.Value))
                cmd.Parameters.AddWithValue("@Email", If(f.Email, DBNull.Value))
                cmd.Parameters.AddWithValue("@Email2", If(f.Email2, DBNull.Value))
                cmd.Parameters.AddWithValue("@Formacao", If(f.Formacao, DBNull.Value))
                cmd.Parameters.AddWithValue("@Cargo", If(f.Cargo, DBNull.Value))
                cmd.Parameters.AddWithValue("@DataAdmissao", If(f.DataAdmissao.HasValue, f.DataAdmissao.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@RegimeContratacao", If(f.RegimeContratacao, DBNull.Value))
                cmd.Parameters.AddWithValue("@Salario", f.Salario)
                cmd.Parameters.AddWithValue("@PrazoExperiencia", If(f.PrazoExperiencia, DBNull.Value))
                cmd.Parameters.AddWithValue("@HorarioTrabalho", If(f.HorarioTrabalho, DBNull.Value))
                cmd.Parameters.AddWithValue("@DescontaVT", If(f.DescontaVT, DBNull.Value))
                cmd.Parameters.AddWithValue("@ObservacoesRH", If(f.ObservacoesRH, DBNull.Value))
                cmd.Parameters.AddWithValue("@RG", If(f.RG, DBNull.Value))
                cmd.Parameters.AddWithValue("@RGDataExpedicao", If(f.RGDataExpedicao.HasValue, f.RGDataExpedicao.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@RGOrgaoExp", If(f.RGOrgaoExp, DBNull.Value))
                cmd.Parameters.AddWithValue("@RGUF", If(f.RGUF, DBNull.Value))
                cmd.Parameters.AddWithValue("@CPF", If(f.CPF, DBNull.Value))
                cmd.Parameters.AddWithValue("@PIS", If(f.PIS, DBNull.Value))
                cmd.Parameters.AddWithValue("@CTPS", If(f.CTPS, DBNull.Value))
                cmd.Parameters.AddWithValue("@CTPSSerie", If(f.CTPSSerie, DBNull.Value))
                cmd.Parameters.AddWithValue("@CTPSUF", If(f.CTPSUF, DBNull.Value))
                cmd.Parameters.AddWithValue("@CTPSDataExp", If(f.CTPSDataExp.HasValue, f.CTPSDataExp.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@TituloEleitor", If(f.TituloEleitor, DBNull.Value))
                cmd.Parameters.AddWithValue("@TituloZona", If(f.TituloZona, DBNull.Value))
                cmd.Parameters.AddWithValue("@TituloSecao", If(f.TituloSecao, DBNull.Value))
                cmd.Parameters.AddWithValue("@TituloMunicipioUF", If(f.TituloMunicipioUF, DBNull.Value))
                cmd.Parameters.AddWithValue("@TituloDataEmissao", If(f.TituloDataEmissao.HasValue, f.TituloDataEmissao.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@Reservista", If(f.Reservista, DBNull.Value))
                cmd.Parameters.AddWithValue("@CNH", If(f.CNH, DBNull.Value))
                cmd.Parameters.AddWithValue("@CNHCategoria", If(f.CNHCategoria, DBNull.Value))
                cmd.Parameters.AddWithValue("@CNHValidade", If(f.CNHValidade.HasValue, f.CNHValidade.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@CNHData1Hab", If(f.CNHData1Hab.HasValue, f.CNHData1Hab.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@CNHDataEmissao", If(f.CNHDataEmissao.HasValue, f.CNHDataEmissao.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@ConjugeDataNasc", If(f.ConjugeDataNasc.HasValue, f.ConjugeDataNasc.Value.ToString("yyyy-MM-dd"), DBNull.Value))
                cmd.Parameters.AddWithValue("@ConjugeRG", If(f.ConjugeRG, DBNull.Value))
                cmd.Parameters.AddWithValue("@ConjugeRGUF", If(f.ConjugeRGUF, DBNull.Value))
                cmd.Parameters.AddWithValue("@ConjugeCPF", If(f.ConjugeCPF, DBNull.Value))
                cmd.Parameters.AddWithValue("@Banco", If(f.Banco, DBNull.Value))
                cmd.Parameters.AddWithValue("@Agencia", If(f.Agencia, DBNull.Value))
                cmd.Parameters.AddWithValue("@ContaCorrente", If(f.ContaCorrente, DBNull.Value))
                cmd.Parameters.AddWithValue("@ContaModalidade", If(f.ContaModalidade, DBNull.Value))
                cmd.Parameters.AddWithValue("@PIX", If(f.PIX, DBNull.Value))
                cmd.Parameters.AddWithValue("@Banco2", If(f.Banco2, DBNull.Value))
                cmd.Parameters.AddWithValue("@Agencia2", If(f.Agencia2, DBNull.Value))
                cmd.Parameters.AddWithValue("@ContaCorrente2", If(f.ContaCorrente2, DBNull.Value))
                cmd.Parameters.AddWithValue("@ContaModalidade2", If(f.ContaModalidade2, DBNull.Value))
                cmd.Parameters.AddWithValue("@PIX2", If(f.PIX2, DBNull.Value))
                cmd.Parameters.AddWithValue("@Dependentes", If(f.Dependentes, DBNull.Value))
                cmd.Parameters.AddWithValue("@EmergenciaNome", If(f.EmergenciaNome, DBNull.Value))
                cmd.Parameters.AddWithValue("@EmergenciaFone", If(f.EmergenciaFone, DBNull.Value))
                cmd.Parameters.AddWithValue("@EmergenciaParentesco", If(f.EmergenciaParentesco, DBNull.Value))
                cmd.Parameters.AddWithValue("@Ativo", f.Ativo)
                If f.REG = 0 AndAlso Not usarRegEspecifico Then
                    Dim novoId = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                    f.REG = novoId
                Else
                    Await cmd.ExecuteNonQueryAsync()
                End If
                Return f.REG
            End Using
        End Function

        Public Async Function AtualizarSituacaoAsync(reg As Integer, situacao As String) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE Funcionario SET Situacao = @s WHERE REG=@reg"
                cmd.Parameters.AddWithValue("@s", situacao)
                cmd.Parameters.AddWithValue("@reg", reg)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function InativarAsync(reg As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE Funcionario SET Ativo = 0 WHERE REG=@reg"
                cmd.Parameters.AddWithValue("@reg", reg)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function ExcluirAsync(reg As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Funcionario WHERE REG=@reg"
                cmd.Parameters.AddWithValue("@reg", reg)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function ExcluirInativosAsync() As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM Funcionario WHERE Ativo = 0"
                Await cmd.ExecuteNonQueryAsync()
                cmd.CommandText = "SELECT COUNT(*) FROM Funcionario"
                Dim qtd = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                If qtd = 0 Then
                    cmd.CommandText = "DBCC CHECKIDENT ('Funcionario', RESEED, 0)"
                    Await cmd.ExecuteNonQueryAsync()
                End If
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 8. REPOSITÓRIO HISTÓRICO VALOR BASE
    ' ==================================================================================
    Public Class HistoricoValorBaseRepo
        Public Async Function ListarAsync(funcionarioREG As Integer) As Task(Of List(Of HistoricoValorBase))
            Dim lst As New List(Of HistoricoValorBase)
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT * FROM HistoricoValorBase WHERE FuncionarioREG=@reg ORDER BY DataAlteracao DESC"
                cmd.Parameters.AddWithValue("@reg", funcionarioREG)
                Using rd = Await cmd.ExecuteReaderAsync()
                    While Await rd.ReadAsync()
                        Dim h As New HistoricoValorBase With {
                            .Id = Convert.ToInt32(rd("Id")),
                            .FuncionarioREG = Convert.ToInt32(rd("FuncionarioREG")),
                            .DataAlteracao = Convert.ToDateTime(rd("DataAlteracao")),
                            .ValorBase = CDec(rd("ValorBase"))
                        }
                        lst.Add(h)
                    End While
                End Using
            End Using
            Return lst
        End Function

        Public Async Function InserirHistoricoAsync(reg As Integer, dataAlt As Date, valor As Decimal) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "INSERT INTO HistoricoValorBase(FuncionarioREG, DataAlteracao, ValorBase) VALUES(@reg, @dt, @v)"
                cmd.Parameters.AddWithValue("@reg", reg)
                cmd.Parameters.AddWithValue("@dt", dataAlt)
                cmd.Parameters.AddWithValue("@v", valor)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function AtualizarHistoricoAsync(id As Integer, dataAlt As Date, valor As Decimal) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE HistoricoValorBase SET DataAlteracao=@dt, ValorBase=@v WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                cmd.Parameters.AddWithValue("@dt", dataAlt)
                cmd.Parameters.AddWithValue("@v", valor)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Async Function ExcluirHistoricoAsync(id As Integer) As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                cmd.CommandText = "DELETE FROM HistoricoValorBase WHERE Id=@id"
                cmd.Parameters.AddWithValue("@id", id)
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Class

    ' ==================================================================================
    ' 9. REPOSITÓRIO USUÁRIO
    ' ==================================================================================
    Public Class UsuarioRepo
        Public Async Function CriarAdminPadraoAsync() As Task
            Using cn = Await Db.GetConnectionAsync(), cmd = cn.CreateCommand()
                ' Insere o usuário admin inicial caso a tabela esteja vazia
                cmd.CommandText = "IF NOT EXISTS (SELECT 1 FROM Usuario WHERE Username = 'admin') " &
                                  "INSERT INTO Usuario (Username, PasswordHash, Salt, Role, MustChangePassword) " &
                                  "VALUES ('admin', 'hash_provisorio', 'salt_provisorio', 'Admin', 1)"
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function

        Public Function ValidarLogin(usuarioLogin As String, senha As String) As FinanceiroPJ.Models.Usuario
            ' OBS: Este método está síncrono para compatibilidade com o AuthService atual.
            ' O ideal no futuro é migrá-lo para Async.
            Dim user As FinanceiroPJ.Models.Usuario = Nothing

            ' Note o uso do .Result para rodar sincronicamente, já que a assinatura não é Async
            Using cn = Db.GetConnectionAsync().GetAwaiter().GetResult()
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "SELECT * FROM Usuario WHERE Username = @usr"
                    cmd.Parameters.AddWithValue("@usr", usuarioLogin)

                    Using rd = cmd.ExecuteReader()
                        If rd.Read() Then
                            user = New FinanceiroPJ.Models.Usuario With {
                                .Username = rd("Username").ToString(),
                                .PasswordHash = rd("PasswordHash").ToString(),
                                .Salt = rd("Salt").ToString(),
                                .Role = rd("Role").ToString(),
                                .MustChangePassword = Convert.ToBoolean(rd("MustChangePassword"))
                            }
                            If Not IsDBNull(rd("ColaboradorId")) Then
                                user.ColaboradorId = Convert.ToInt32(rd("ColaboradorId"))
                            End If
                        End If
                    End Using
                End Using
            End Using

            ' TODO: Aqui você deve implementar a verificação real do Hash da senha.
            ' Se a validação do hash bater, retorne o user; senão, retorne Nothing.
            Return user
        End Function
    End Class

End Namespace