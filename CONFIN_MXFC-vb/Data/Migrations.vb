'======================================
' Migrations.vb
' VERSÃO: 15.0.0 (Migração Web - 100% Async)
'=======================================
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports System

Namespace Data
    Public Module Migrations
        Private Const CURRENT_DB_VERSION As Integer = 14

        Public Async Function EnsureCreatedAsync() As Task
            Using cn = Await Db.GetConnectionAsync()
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AppConfig') " &
                                      "CREATE TABLE AppConfig (Chave VARCHAR(50) PRIMARY KEY, Valor VARCHAR(MAX));"
                    Await cmd.ExecuteNonQueryAsync()
                End Using

                Dim dbVersion As Integer = 0
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "SELECT Valor FROM AppConfig WHERE Chave = 'DB_VERSION'"
                    Dim result = Await cmd.ExecuteScalarAsync()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        dbVersion = Convert.ToInt32(result)
                    End If
                End Using

                If dbVersion < CURRENT_DB_VERSION Then
                    Await ExecutarMigracoesAsync(cn, dbVersion, CURRENT_DB_VERSION)
                End If

                Await CriarTabelasBaseAsync(cn)
            End Using

            Dim userRepo As New UsuarioRepo()
            Await userRepo.CriarAdminPadraoAsync()
        End Function

        Private Async Function ExecutarMigracoesAsync(cn As SqlConnection, versaoAtual As Integer, versaoAlvo As Integer) As Task
            For v = versaoAtual + 1 To versaoAlvo
                Using cmd = cn.CreateCommand()
                    Select Case v
                        Case 2, 3, 4, 5, 6, 7, 8
                        Case 9
                            cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Cliente') " &
                                              "CREATE TABLE Cliente (" &
                                              " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                              " RazaoSocial VARCHAR(MAX) NOT NULL, " &
                                              " CNPJ VARCHAR(50), " &
                                              " InscricaoEstadual VARCHAR(50), " &
                                              " Endereco VARCHAR(MAX), " &
                                              " CEP VARCHAR(20), " &
                                              " Telefone VARCHAR(50), " &
                                              " Responsavel VARCHAR(200));"
                            Await cmd.ExecuteNonQueryAsync()
                        Case 10
                            cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Contrato') " &
                                              "CREATE TABLE Contrato (" &
                                              " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                              " ClienteId INT NOT NULL, " &
                                              " NumeroContrato VARCHAR(100) NOT NULL, " &
                                              " Valor DECIMAL(18,2) NOT NULL, " &
                                              " Data DATETIME NOT NULL, " &
                                              " CentroCusto VARCHAR(100), " &
                                              " Anexo VARBINARY(MAX), " &
                                              " FOREIGN KEY(ClienteId) REFERENCES Cliente(Id) ON DELETE CASCADE);"
                            Await cmd.ExecuteNonQueryAsync()
                        Case 11
                            cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ContratoAditivo') " &
                                              "CREATE TABLE ContratoAditivo (" &
                                              " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                              " ContratoId INT NOT NULL, " &
                                              " Numero VARCHAR(100), " &
                                              " Data DATETIME NOT NULL, " &
                                              " Valor DECIMAL(18,2) DEFAULT 0, " &
                                              " Descricao VARCHAR(MAX), " &
                                              " Anexo VARBINARY(MAX), " &
                                              " FOREIGN KEY(ContratoId) REFERENCES Contrato(Id) ON DELETE CASCADE);"
                            Await cmd.ExecuteNonQueryAsync()
                        Case 12
                            Try
                                cmd.CommandText = "ALTER TABLE Contrato ADD DataInicial DATE, DataFinal DATE, Prazo INT DEFAULT 0;"
                                Await cmd.ExecuteNonQueryAsync()
                                cmd.CommandText = "ALTER TABLE ContratoAditivo ADD DataFinal DATE;"
                                Await cmd.ExecuteNonQueryAsync()
                                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ContratoAnexo') " &
                                                  "CREATE TABLE ContratoAnexo (" &
                                                  " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                                  " ContratoId INT NOT NULL, " &
                                                  " NomeArquivo VARCHAR(255), " &
                                                  " Conteudo VARBINARY(MAX), " &
                                                  " FOREIGN KEY(ContratoId) REFERENCES Contrato(Id) ON DELETE CASCADE);"
                                Await cmd.ExecuteNonQueryAsync()
                            Catch
                            End Try
                        Case 13
                            Try
                                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'ColaboradorId' AND Object_ID = Object_ID(N'Usuario')) " &
                                                  "BEGIN " &
                                                  "    ALTER TABLE Usuario ADD ColaboradorId INT NULL; " &
                                                  "END"
                                Await cmd.ExecuteNonQueryAsync()

                                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'FK_Usuario_Funcionario') " &
                                                  "BEGIN " &
                                                  "    ALTER TABLE Usuario ADD CONSTRAINT FK_Usuario_Funcionario FOREIGN KEY (ColaboradorId) REFERENCES Funcionario(REG); " &
                                                  "END"
                                Await cmd.ExecuteNonQueryAsync()
                            Catch ex As Exception
                                Throw New Exception("Erro na Migração 13 (Vínculo Usuário): " & ex.Message)
                            End Try
                        Case 14
                            Try
                                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'HistoricoValorBase') " &
                                                  "CREATE TABLE HistoricoValorBase (" &
                                                  " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                                  " FuncionarioREG INT NOT NULL, " &
                                                  " DataAlteracao DATE NOT NULL, " &
                                                  " ValorBase DECIMAL(18,2) NOT NULL, " &
                                                  " CONSTRAINT FK_Historico_Funcionario FOREIGN KEY (FuncionarioREG) REFERENCES Funcionario(REG) ON DELETE CASCADE);"
                                Await cmd.ExecuteNonQueryAsync()
                            Catch ex As Exception
                                Throw New Exception("Erro na Migração 14 (HistoricoValorBase): " & ex.Message)
                            End Try
                    End Select

                    cmd.CommandText = "UPDATE AppConfig SET Valor = @v WHERE Chave = 'DB_VERSION'; " &
                                      "IF @@ROWCOUNT = 0 INSERT INTO AppConfig (Chave, Valor) VALUES ('DB_VERSION', @v);"
                    cmd.Parameters.AddWithValue("@v", v.ToString())
                    Await cmd.ExecuteNonQueryAsync()
                End Using
            Next
        End Function

        Private Async Function CriarTabelasBaseAsync(cn As SqlConnection) As Task
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Funcionario') " &
                "CREATE TABLE Funcionario (" &
                " REG INT IDENTITY(1,1) PRIMARY KEY, " &
                " NomeCompleto VARCHAR(MAX) NOT NULL, " &
                " Foto VARBINARY(MAX), " &
                " Declaracao VARBINARY(MAX), " &
                " Situacao VARCHAR(50) DEFAULT 'Ativo', DataNascimento VARCHAR(20), Naturalidade VARCHAR(100), Nacionalidade VARCHAR(100), " &
                " EstadoCivil VARCHAR(50), NomeConjuge VARCHAR(200), NomePai VARCHAR(200), NomeMae VARCHAR(200), " &
                " Endereco VARCHAR(MAX), EnderecoNumero VARCHAR(50), EnderecoComplemento VARCHAR(100), EnderecoBairro VARCHAR(100), " &
                " CEP VARCHAR(20), Cidade VARCHAR(100), Estado VARCHAR(50), Telefone VARCHAR(50), Celular VARCHAR(50), " &
                " Email VARCHAR(200), Email2 VARCHAR(200), Formacao VARCHAR(MAX), Cargo VARCHAR(100), " &
                " DataAdmissao VARCHAR(20), RegimeContratacao VARCHAR(100), " &
                " Salario DECIMAL(18,2) DEFAULT 0, " &
                " PrazoExperiencia VARCHAR(100), HorarioTrabalho VARCHAR(MAX), DescontaVT VARCHAR(10), ObservacoesRH VARCHAR(MAX), " &
                " RG VARCHAR(50), RGDataExpedicao VARCHAR(20), RGOrgaoExp VARCHAR(50), RGUF VARCHAR(10), " &
                " CPF VARCHAR(50), PIS VARCHAR(50), CTPS VARCHAR(50), CTPSSerie VARCHAR(50), CTPSUF VARCHAR(10), CTPSDataExp VARCHAR(20), " &
                " TituloEleitor VARCHAR(50), TituloZona VARCHAR(20), TituloSecao VARCHAR(20), TituloMunicipioUF VARCHAR(50), TituloDataEmissao VARCHAR(20), " &
                " Reservista VARCHAR(50), CNH VARCHAR(50), CNHCategoria VARCHAR(10), CNHValidade VARCHAR(20), CNHData1Hab VARCHAR(20), CNHDataEmissao VARCHAR(20), " &
                " ConjugeDataNasc VARCHAR(20), ConjugeRG VARCHAR(50), ConjugeRGUF VARCHAR(10), ConjugeCPF VARCHAR(50), " &
                " Banco VARCHAR(100), Agencia VARCHAR(50), ContaCorrente VARCHAR(50), ContaModalidade VARCHAR(50), PIX VARCHAR(100), " &
                " Banco2 VARCHAR(100), Agencia2 VARCHAR(50), ContaCorrente2 VARCHAR(50), ContaModalidade2 VARCHAR(50), PIX2 VARCHAR(100), " &
                " Dependentes VARCHAR(MAX), EmergenciaNome VARCHAR(200), EmergenciaFone VARCHAR(50), EmergenciaParentesco VARCHAR(50), " &
                " Ativo BIT DEFAULT 1);"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Colaborador') " &
                "CREATE TABLE Colaborador (" &
                " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                " NomeEmpresa VARCHAR(MAX) NOT NULL, CNPJ VARCHAR(50), Endereco VARCHAR(MAX), Cidade VARCHAR(100), CEP VARCHAR(20), " &
                " Responsavel VARCHAR(200), Telefone VARCHAR(50), Email VARCHAR(200), PixOuDadosBancarios VARCHAR(MAX), " &
                " FuncionarioREG INT, " &
                " FOREIGN KEY(FuncionarioREG) REFERENCES Funcionario(REG) ON DELETE CASCADE);"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Cliente') " &
                "CREATE TABLE Cliente (" &
                " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                " RazaoSocial VARCHAR(MAX) NOT NULL, " &
                " CNPJ VARCHAR(50), " &
                " InscricaoEstadual VARCHAR(50), " &
                " Endereco VARCHAR(MAX), " &
                " CEP VARCHAR(20), " &
                " Telefone VARCHAR(50), " &
                " Responsavel VARCHAR(200));"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Gasto') " &
                "CREATE TABLE Gasto (" &
                " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                " ColaboradorId INT NOT NULL, " &
                " Data DATETIME NOT NULL, " &
                " NomeDespesa VARCHAR(MAX) NOT NULL, TipoDespesa VARCHAR(100) NOT NULL, " &
                " Valor DECIMAL(18,2) NOT NULL, " &
                " Anexo VARBINARY(MAX), " &
                " FOREIGN KEY(ColaboradorId) REFERENCES Colaborador(Id) ON DELETE CASCADE);"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'NotaFiscal') " &
                "CREATE TABLE NotaFiscal (" &
                " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                " ColaboradorId INT NOT NULL, Mes INT NOT NULL, Ano INT NOT NULL, " &
                " ValorFixo DECIMAL(18,2) DEFAULT 0, FG DECIMAL(18,2) DEFAULT 0, Adicional DECIMAL(18,2) DEFAULT 0, " &
                " CentroCusto VARCHAR(100), Despesas DECIMAL(18,2) DEFAULT 0, Contador DECIMAL(18,2) DEFAULT 0, " &
                " INSSProlabore DECIMAL(18,2) DEFAULT 0, ImpostoNF DECIMAL(18,2) DEFAULT 0, ValorNF DECIMAL(18,2) DEFAULT 0, " &
                " Prolabore DECIMAL(18,2) DEFAULT 0, Observacao VARCHAR(MAX), " &
                " CONSTRAINT UK_NotaFiscal UNIQUE(ColaboradorId, Mes, Ano), " &
                " FOREIGN KEY(ColaboradorId) REFERENCES Colaborador(Id) ON DELETE CASCADE);"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Usuario') " &
                "CREATE TABLE Usuario (" &
                " Username VARCHAR(50) PRIMARY KEY NOT NULL, " &
                " PasswordHash VARCHAR(MAX) NOT NULL, Salt VARCHAR(MAX) NOT NULL, " &
                " Role VARCHAR(50) NOT NULL, MustChangePassword BIT DEFAULT 0, " &
                " ColaboradorId INT NULL, " &
                " FOREIGN KEY(ColaboradorId) REFERENCES Funcionario(REG));"
                Await cmd.ExecuteNonQueryAsync()

                cmd.CommandText = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'HistoricoValorBase') " &
                                  "CREATE TABLE HistoricoValorBase (" &
                                  " Id INT IDENTITY(1,1) PRIMARY KEY, " &
                                  " FuncionarioREG INT NOT NULL, " &
                                  " DataAlteracao DATE NOT NULL, " &
                                  " ValorBase DECIMAL(18,2) NOT NULL, " &
                                  " CONSTRAINT FK_Historico_Funcionario FOREIGN KEY (FuncionarioREG) REFERENCES Funcionario(REG) ON DELETE CASCADE);"
                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Function
    End Module
End Namespace