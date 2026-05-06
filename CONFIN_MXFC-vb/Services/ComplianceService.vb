'======================================
' ComplianceService.vb
' VERSÃO: 15.0.0 (Exclusivo Web - MemoryStream)
'=======================================
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports CONFIN_MXFC.FinanceiroPJ.Models
Imports System.Text.RegularExpressions
Imports System

Namespace Services
    Public Class ComplianceService

        Private Shared Sub MontarDocumento(doc As Document, writer As PdfWriter, f As Funcionario, imgFundoPath As String)
            If File.Exists(imgFundoPath) Then
                writer.PageEvent = New BackgroundPageEvent(imgFundoPath)
            End If

            doc.Open()

            Dim fontTitulo As Font = FontFactory.GetFont("Arial", 14, Font.BOLD, BaseColor.BLACK)
            Dim fontSubTitulo As Font = FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.BLACK)
            Dim fontNormal As Font = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLACK)
            Dim fontBold As Font = FontFactory.GetFont("Arial", 12, Font.BOLD, BaseColor.BLACK)

            Dim pTitulo As New Paragraph("DECLARAÇÃO DE CONHECIMENTO", fontTitulo) With {
                .Alignment = Element.ALIGN_CENTER,
                .SpacingAfter = 60
            }
            doc.Add(pTitulo)

            Dim pSubTitulo As New Paragraph("COMPLIANCE MXFC ENGENHARIA", fontSubTitulo) With {
                .Alignment = Element.ALIGN_RIGHT,
                .SpacingAfter = 50
            }
            doc.Add(pSubTitulo)

            Dim cpfFormatado As String = FormatarCPF(f.CPF)
            Dim p As New Paragraph() With {
                .Alignment = Element.ALIGN_JUSTIFIED
            }
            p.SetLeading(0, 1.5F)
            p.Add(New Chunk("Eu, ", fontNormal))
            p.Add(New Chunk(f.NomeCompleto.ToUpper(), fontBold))
            p.Add(New Chunk(", CPF ", fontNormal))
            p.Add(New Chunk(cpfFormatado, fontBold))
            p.Add(New Chunk(", cargo ", fontNormal))
            Dim cargoTexto = If(String.IsNullOrEmpty(f.Cargo), "________________", f.Cargo.ToUpper())
            p.Add(New Chunk(cargoTexto, fontBold))
            p.Add(New Chunk(", declaro que li e compreendi o Código de Ética e Conduta da MXFC e me comprometo a cumpri-lo, comunicando ao Compliance qualquer situação de risco ou violação.", fontNormal))
            doc.Add(p)

            doc.Add(New Paragraph(vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf))

            Dim dataStr As String = "____ / ____ / _______"
            If f.DataAdmissao.HasValue Then
                dataStr = f.DataAdmissao.Value.ToString("dd / MM / yyyy")
            End If
            Dim pData As New Paragraph($"Data:  {dataStr}", fontBold) With {
                .Alignment = Element.ALIGN_RIGHT,
                .SpacingAfter = 50
            }
            doc.Add(pData)

            Dim line As New iTextSharp.text.pdf.draw.LineSeparator(1.0F, 60.0F, BaseColor.BLACK, Element.ALIGN_CENTER, -1)
            doc.Add(New Chunk(line))

            Dim pNome As New Paragraph(f.NomeCompleto.ToUpper(), fontBold) With {
                .Alignment = Element.ALIGN_CENTER
            }
            doc.Add(pNome)
        End Sub

        Public Shared Function GerarTermoCompromissoBytes(f As Funcionario, imgFundoPath As String) As Byte()
            Dim doc As New Document(PageSize.A4, 71, 57, 150, 50)
            Using ms As New MemoryStream()
                Try
                    Dim writer = PdfWriter.GetInstance(doc, ms)
                    MontarDocumento(doc, writer, f, imgFundoPath)
                    doc.Close()
                    Return ms.ToArray()
                Catch ex As Exception
                    Throw New Exception("Erro ao gerar PDF em memória: " & ex.Message)
                End Try
            End Using
        End Function

        Private Shared Function FormatarCPF(cpf As String) As String
            If String.IsNullOrWhiteSpace(cpf) Then Return "___.___.___-__"
            Dim apenasNumeros = Regex.Replace(cpf, "[^\d]", "")
            If apenasNumeros.Length = 11 Then
                Return Convert.ToInt64(apenasNumeros).ToString("000\.000\.000\-00")
            End If
            Return cpf
        End Function

        Private Class BackgroundPageEvent
            Inherits PdfPageEventHelper
            Private ReadOnly _img As Image
            Public Sub New(imgPath As String)
                _img = Image.GetInstance(imgPath)
                _img.ScaleToFit(PageSize.A4.Width, PageSize.A4.Height)
                _img.SetAbsolutePosition(0, 0)
            End Sub
            Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
                writer.DirectContentUnder.AddImage(_img)
            End Sub
        End Class
    End Class
End Namespace