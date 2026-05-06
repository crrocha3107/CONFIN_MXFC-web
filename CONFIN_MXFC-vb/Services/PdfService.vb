'======================================
' PdfService.vb
' VERSÃO: 15.0.0 (Exclusivo Web)
'=======================================
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Data
Imports System.Linq
Imports System

Namespace Services
    Public Class PdfService
        Private Const MARGEM_SUP_CM As Single = 3.0F
        Private Const MARGEM_INF_CM As Single = 2.0F
        Private Const MARGEM_ESQ_CM As Single = 2.5F
        Private Const MARGEM_DIR_CM As Single = 2.0F
        Private Const CM_TO_POINTS As Single = 28.34645F

        Public Shared Function ExportarTabelaParaPdfBytes(dt As DataTable, titulo As String, Optional caminhoImagemFundo As String = "") As Byte()
            Dim doc As New Document(PageSize.A4,
                                    MARGEM_ESQ_CM * CM_TO_POINTS,
                                    MARGEM_DIR_CM * CM_TO_POINTS,
                                    MARGEM_SUP_CM * CM_TO_POINTS,
                                    MARGEM_INF_CM * CM_TO_POINTS)

            Using ms As New MemoryStream()
                Try
                    Dim writer = PdfWriter.GetInstance(doc, ms)

                    If Not String.IsNullOrWhiteSpace(caminhoImagemFundo) Then
                        If File.Exists(caminhoImagemFundo) Then
                            writer.PageEvent = New BackgroundPageEvent(caminhoImagemFundo)
                        End If
                    End If

                    doc.Open()

                    Dim fonteTitulo As New Font(Font.FontFamily.HELVETICA, 16, Font.BOLD)
                    Dim paragrafoTitulo As New Paragraph(titulo, fonteTitulo) With {
                        .Alignment = Element.ALIGN_CENTER,
                        .SpacingAfter = 20
                    }
                    doc.Add(paragrafoTitulo)

                    Dim tabela As New PdfPTable(dt.Columns.Count) With {
                        .WidthPercentage = 100
                    }

                    Dim fonteHeader As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD, BaseColor.WHITE)
                    For Each col As DataColumn In dt.Columns
                        Dim cell As New PdfPCell(New Phrase(col.ColumnName, fonteHeader)) With {
                            .BackgroundColor = New BaseColor(50, 50, 50),
                            .HorizontalAlignment = Element.ALIGN_CENTER,
                            .Padding = 5
                        }
                        tabela.AddCell(cell)
                    Next

                    Dim fonteDados As New Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL)
                    For Each row As DataRow In dt.Rows
                        For Each col As DataColumn In dt.Columns
                            Dim valor = If(row(col), "").ToString()
                            Dim cell As New PdfPCell(New Phrase(valor, fonteDados)) With {
                                .Padding = 5,
                                .HorizontalAlignment = If(IsNumeric(valor.Replace("R$", "").Trim()), Element.ALIGN_RIGHT, Element.ALIGN_LEFT)
                            }
                            tabela.AddCell(cell)
                        Next
                    Next

                    doc.Add(tabela)
                    doc.Close()

                    Return ms.ToArray()
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End Using
        End Function

        Private Class BackgroundPageEvent
            Inherits PdfPageEventHelper
            Private ReadOnly _imgPath As String
            Private ReadOnly _img As Image

            Public Sub New(imgPath As String)
                _imgPath = imgPath
                _img = Image.GetInstance(_imgPath)
                _img.ScaleToFit(PageSize.A4.Width, PageSize.A4.Height)
                _img.SetAbsolutePosition(0, 0)
            End Sub

            Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
                Dim content As PdfContentByte = writer.DirectContentUnder
                content.AddImage(_img)
            End Sub
        End Class
    End Class
End Namespace