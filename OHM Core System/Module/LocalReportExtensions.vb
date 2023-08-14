Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.IO
Imports System.Runtime.CompilerServices

Module LocalReportExtensions
    <Extension()>
    Sub PrintToPrinter(ByVal report As LocalReport)
        Dim pageSettings As PageSettings = New PageSettings()
        pageSettings.PaperSize = report.GetDefaultPageSettings().PaperSize
        pageSettings.Landscape = report.GetDefaultPageSettings().IsLandscape
        pageSettings.Margins = report.GetDefaultPageSettings().Margins
        Print(report, pageSettings)
    End Sub

    <Extension()>
    Sub Print(ByVal report As LocalReport, ByVal pageSettings As PageSettings)
        Dim deviceInfo As String = $"<DeviceInfo>
                    <OutputFormat>EMF</OutputFormat>
                    <PageWidth>{pageSettings.PaperSize.Width * 100}in</PageWidth>
                    <PageHeight>{pageSettings.PaperSize.Height * 100}in</PageHeight>
                    <MarginTop>{pageSettings.Margins.Top * 100}in</MarginTop>
                    <MarginLeft>{pageSettings.Margins.Left * 100}in</MarginLeft>
                    <MarginRight>{pageSettings.Margins.Right * 100}in</MarginRight>
                    <MarginBottom>{pageSettings.Margins.Bottom * 100}in</MarginBottom>
                </DeviceInfo>"
        Dim warnings As Warning()
        Dim streams = New List(Of Stream)()
        Dim pageIndex = 0
        report.Render("Image", deviceInfo, Function(name, fileNameExtension, encoding, mimeType, willSeek)
                                               Dim stream As MemoryStream = New MemoryStream()
                                               streams.Add(stream)
                                               Return stream
                                           End Function, warnings)

        For Each stream As Stream In streams
            stream.Position = 0
        Next

        If streams Is Nothing OrElse streams.Count = 0 Then Throw New Exception("No stream to print.")

        Using printDocument As PrintDocument = New PrintDocument()
            printDocument.DefaultPageSettings = pageSettings

            If Not printDocument.PrinterSettings.IsValid Then
                Throw New Exception("Can't find the default printer.")
            Else
                printDocument.PrintPage += Function(sender, e)
                                               Dim pageImage As Metafile = New Metafile(streams(pageIndex))
                                               Dim adjustedRect As Rectangle = New Rectangle(e.PageBounds.Left - CInt(e.PageSettings.HardMarginX), e.PageBounds.Top - CInt(e.PageSettings.HardMarginY), e.PageBounds.Width, e.PageBounds.Height)
                                               e.Graphics.FillRectangle(Brushes.White, adjustedRect)
                                               e.Graphics.DrawImage(pageImage, adjustedRect)
                                               pageIndex += 1
                                               e.HasMorePages = (pageIndex < streams.Count)
                                               e.Graphics.DrawRectangle(Pens.Red, adjustedRect)
                                           End Function

                printDocument.EndPrint += Function(Sender, e)

                                              If streams IsNot Nothing Then

                                                  For Each stream As Stream In streams
                                                      stream.Close()
                                                  Next

                                                  streams = Nothing
                                              End If
                                          End Function

                printDocument.Print()
            End If
        End Using
    End Sub
End Module
