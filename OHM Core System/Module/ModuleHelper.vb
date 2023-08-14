Imports ClosedXML.Excel
Imports System.Text.RegularExpressions
Module ModuleHelper
    Sub ExportDgvToExcel(ByVal dgv As DataGridView)
        Try
            If dgv Is Nothing OrElse dgv.RowCount <= 0 Then
                MsgBox("The Data is empty!!!")
                Exit Sub
            End If
            Dim sfdExcel As SaveFileDialog = New SaveFileDialog
            sfdExcel.Filter = "Excel File|*.xlsx"
            sfdExcel.Title = "Save an Excel File"
            sfdExcel.ShowDialog()
            If sfdExcel.FileName = "" Then
                MsgBox("Please Type File Name")
                Exit Sub
            End If
            Dim wb As XLWorkbook = New XLWorkbook
            wb.AddWorksheet(Format(Now, "dd MMM yy"))
            Dim ws = wb.Worksheet(Format(Now, "dd MMM yy"))
            For Each column As DataGridViewColumn In dgv.Columns
                ws.Cell(1, column.Index + 1).Value = column.HeaderText
                With ws.Cell(1, column.Index + 1).Style
                    .Font.Bold = True
                    .Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    .Border.OutsideBorder = XLBorderStyleValues.Thin
                    .Fill.BackgroundColor = XLColor.LightGreen
                End With
                ws.RowHeight = dgv.ColumnHeadersHeight
            Next
            For Each row As DataGridViewRow In dgv.Rows
                For Each cell As DataGridViewCell In row.Cells
                    ws.Cell(cell.RowIndex + 2, cell.ColumnIndex + 1).Value = cell.Value.ToString
                Next
            Next
            wb.SaveAs(sfdExcel.FileName)
            If MsgBox("Do You Want To Open File", vbQuestion + vbYesNo, "Question?") = vbYes Then
                Process.Start(sfdExcel.FileName)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'WriteErrorLog("sub saveExcelFile Form7--> Error saat export data")
        End Try
    End Sub

    Sub ImportExcelToDgv(ByVal dgv As DataGridView, sheetName As String, startCell As String, endCell As String)
        Try
            Dim workbook As XLWorkbook
            Dim ofd As New OpenFileDialog

            ofd.Filter = "Excel files (*.xlsx, *.xls) | *.xlsx; *.xls"
            If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                workbook = New XLWorkbook(ofd.FileName)
                Dim wsheet = workbook.Worksheet(sheetName)

                Dim startRow = CInt(System.Text.RegularExpressions.Regex.Replace(startCell, "[^\d]", ""))
                Dim endRow = CInt(System.Text.RegularExpressions.Regex.Replace(endCell, "[^\d]", ""))
                Dim startColumn = wsheet.Row(startRow).Cell(Regex.Replace(startCell, "\d+$", "")).Address.ColumnNumber - 1
                Dim endColumn = wsheet.Row(startRow).Cell(Regex.Replace(endCell, "\d+$", "")).Address.ColumnNumber - 1

                '------Get Value Data
                Dim dt As New DataTable()
                For i = startColumn To endColumn
                    Dim dc As New DataColumn()
                    dc.ColumnName = wsheet.Rows(startRow).Cells(i).Value.ToString
                    dc.DataType = GetType(String)
                    dt.Columns.Add(dc)
                Next
                Dim ctr = startRow + 1
                While ctr <> endRow
                    Dim dr As DataRow = dt.NewRow
                    For i = startColumn To endColumn
                        dr(i) = wsheet.Rows(ctr).Cells(i).Value
                    Next
                    dt.Rows.Add(dr)
                    ctr = ctr + 1
                End While
                dgv.DataSource = dt
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub
End Module