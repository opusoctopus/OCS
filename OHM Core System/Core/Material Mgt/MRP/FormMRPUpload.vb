Imports MySql.Data.MySqlClient
Imports System.IO
Imports ExcelDataReader

Public Class FormMRPUpload

    ' CREATE EXCEL OBJECTS.
    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim sFileName As String = ""
    Dim tables As DataTableCollection

    Sub Bersih()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Refresh()
        Label1.Text = ""
    End Sub

    Sub customdgv2()
        With DataGridView1 'Ganti dengan nama datagridview kalian
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 35
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'With OpenFileDialog1
        '.Title = "Excel File to Edit"           ' DIALOG BOX TITLE.
        '.FileName = ""
        '.Filter = "Excel File|*.xlsx;*.xls"     ' FILTER ONLY EXCEL FILES IN FILE TYPE.

        'If .ShowDialog() = DialogResult.OK Then
        'sFileName = .FileName

        'If Trim(sFileName) <> "" Then
        'Excel2Grid(sFileName)
        'End If
        'End If
        'End With

        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel Workbook|*.xlsx"}
            If ofd.ShowDialog() = DialogResult.OK Then
                TextBox1.Text = ofd.FileName
                Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                     .UseHeaderRow = True}})
                        tables = result.Tables
                        ComboBox1.Items.Clear()
                        For Each table As DataTable In tables
                            ComboBox1.Items.Add(table.TableName)
                        Next
                    End Using
                End Using
            End If
        End Using
    End Sub

    ' IMPORT EXCEL DATA TO DATAGRIDVIEW.
    Private Sub Excel2Grid(ByVal sFile As String)
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(sFile)            ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")       ' THE SHEET WITH THE DATA.

        With DataGridView1
            .Rows.Clear()
            .Columns.Clear()
        End With

        ' FIRST, CREATE THE DataGridView COLUMN HEADERS.
        Dim iCol As Integer
        For iCol = 1 To xlWorkSheet.Columns.Count
            If Trim(xlWorkSheet.Cells(1, iCol).value) = "" Then
                Exit For        ' BAIL OUT IF REACHED THE LAST COL.
            Else
                Dim col = New DataGridViewTextBoxColumn()
                col.HeaderText = xlWorkSheet.Cells(1, iCol).value
                Dim colIndex As Integer = DataGridView1.Columns.Add(col)    ' ADD A NEW COLUMN.
            End If
        Next

        ' ADD A BUTTON AT THE LAST COLUMN IN EVERY ROW.
        'Dim btn = New DataGridViewButtonColumn()
        'btn.HeaderText = ""
        'btn.Text = "Save Data"
        'btn.Name = "btSave"
        'btn.UseColumnTextForButtonValue = True
        'DataGridView1.Columns.Add(btn)

        ' ADD ROWS TO THE GRID USING EXCEL DATA.
        Dim iRow As Integer
        For iRow = 2 To xlWorkSheet.Rows.Count
            If Trim(xlWorkSheet.Cells(iRow, 1).value) = "" Then
                Exit For        ' BAIL OUT IF REACHED THE LAST ROW.
            Else
                ' CREATE A STRING ARRAY USING THE VALUES IN EACH ROW OF THE SHEET.
                Dim row As String() = New String() {
                    xlWorkSheet.Cells(iRow, 1).value,
                    xlWorkSheet.Cells(iRow, 2).value.ToString(),
                    xlWorkSheet.Cells(iRow, 3).value,
                    xlWorkSheet.Cells(iRow, 4).value.ToString(),
                    xlWorkSheet.Cells(iRow, 5).value,
                    xlWorkSheet.Cells(iRow, 6).value,
                    xlWorkSheet.Cells(iRow, 7).value,
                    xlWorkSheet.Cells(iRow, 8).value,
                    xlWorkSheet.Cells(iRow, 9).value,
                    xlWorkSheet.Cells(iRow, 10).value,
                    xlWorkSheet.Cells(iRow, 11).value,
                    xlWorkSheet.Cells(iRow, 12).value,
                    xlWorkSheet.Cells(iRow, 13).value,
                    xlWorkSheet.Cells(iRow, 14).value,
                    xlWorkSheet.Cells(iRow, 15).value,
                    xlWorkSheet.Cells(iRow, 16).value,
                    xlWorkSheet.Cells(iRow, 17).value,
                    xlWorkSheet.Cells(iRow, 18).value,
                    xlWorkSheet.Cells(iRow, 19).value,
                    xlWorkSheet.Cells(iRow, 20).value,
                    xlWorkSheet.Cells(iRow, 21).value,
                    xlWorkSheet.Cells(iRow, 22).value,
                    xlWorkSheet.Cells(iRow, 23).value,
                    xlWorkSheet.Cells(iRow, 24).value,
                    xlWorkSheet.Cells(iRow, 25).value,
                    xlWorkSheet.Cells(iRow, 26).value,
                    xlWorkSheet.Cells(iRow, 27).value,
                    xlWorkSheet.Cells(iRow, 28).value,
                    xlWorkSheet.Cells(iRow, 29).value,
                    xlWorkSheet.Cells(iRow, 30).value}

                ' ADD A NEW ROW TO THE GRID USING THE ARRAY DATA.
                DataGridView1.Rows.Add(row)
            End If
        Next

        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
            ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If keyData = Keys.Enter Then
            SendKeys.Send("{TAB}")      ' MOVE NEXT CELL WHEN YOU PRESS ENTER KEY.
            Return True
        Else
            Return MyBase.ProcessCmdKey(msg, keyData)
        End If
    End Function

    ' SAVE MODIFIED OR NEW DATA TO THE EXCEL SHEET.
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' EVERY ROW HAS A BUTTON AT THE LAST COLUMN. 
        ' SAVE THE DATA IN EXCEL AFTER CLICKING THE BUTTON.

        Dim ourGrid = DirectCast(sender, DataGridView)

        If TypeOf ourGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso e.RowIndex >= 0 Then

            xlApp = New Microsoft.Office.Interop.Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(sFileName)
            xlWorkSheet = xlWorkBook.Worksheets("Sheet1")

            ' CHECK IF THE FIRST COLUMN IS ReadOnly. 
            ' THIS IS TO ENSURE THAT YOU MODIFY EXISTING DATA IN EXCEL.
            If DataGridView1.Rows(e.RowIndex).Cells(0).ReadOnly = True Then
                If Trim(xlWorkSheet.Cells(e.RowIndex + 2, 1).value) = DataGridView1.Rows(e.RowIndex).Cells(0).Value Then
                    ' MODIFY THE DATA.
                    xlWorkSheet.Cells(e.RowIndex + 2, 2).value =
                        DataGridView1.Rows(e.RowIndex).Cells(1).Value   ' FIRST COLUMN.
                    xlWorkSheet.Cells(e.RowIndex + 2, 3).value =
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value   ' SECOND COLUMN.
                End If
            Else
                ' ADD NEW EMPLOYEE DATA IN A NEW ROW IN EXCEL.
                xlWorkSheet.Cells(e.RowIndex + 2, 1).value =
                    DataGridView1.Rows(e.RowIndex).Cells(0).Value
                xlWorkSheet.Cells(e.RowIndex + 2, 2).value =
                    DataGridView1.Rows(e.RowIndex).Cells(1).Value
                xlWorkSheet.Cells(e.RowIndex + 2, 3).value =
                    DataGridView1.Rows(e.RowIndex).Cells(2).Value
            End If

            xlWorkBook.Close() : xlApp.Quit()

            ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        End If
    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged
        ' MAKE THE FIRST CELL (WITH EMPLOYEE NAME) READ ONLY.
        If Trim(e.Row.Cells(0).Value) <> "" Then
            e.Row.Cells(0).Style.ForeColor = Color.Gray
            e.Row.Cells(0).ReadOnly = True
        End If
    End Sub

    Private Sub FormMRPUpload_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
    End Sub

    Sub DeleteData()
        Try
            Call connection()
            str = "DELETE from prodsys2_master_mrp_tbl"
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub SaveData()
        'DELETE OLD DATA
        Call DeleteData()

        ' SAVE CURRENT DATA
        For baris As Integer = 0 To DataGridView1.Rows.Count - 1
            cmd = New MySqlCommand("INSERT INTO `prodsys2_master_mrp_tbl`(`org`, `supply_wo`, `pn`, `start_date`, `start_qty`, `complete_qty`, `status`, 
`uom1`, `child`, `level`, `supply_type`, `description`, `planner`, `wh`, `sum_arrive`, `sum_departure`, `deliv_type`, `req_qty`, `issued_qty`, `remain_qty`, `qpa`, `uit`, 
`uom2`, `raw_mtl`, `raw_wip`, `wip_assy`, `item_cost`, `line_code`, `substitute_item`, `created`) 

VALUES (@org, @supply_wo, @pn, @start_date, @start_qty, @complete_qty, @status, @uom1, @child, @level, @supply_type, @description, @planner, @wh
, @sum_arrive, @sum_departure, @deliv_type, @req_qty, @issued_qty, @remain_qty, @qpa, @uit, @uom2, @raw_mtl, @raw_wip, @wip_assy, @item_cost
, @line_code, @substitute_item, @created)", conn)
            cmd.Parameters.Add("@org", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(1).Value.ToString()
            cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(2).Value.ToString()
            cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(3).Value.ToString()
            cmd.Parameters.Add("@start_date", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(4).Value.ToString()
            cmd.Parameters.Add("@start_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(5).Value.ToString()
            cmd.Parameters.Add("@complete_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(6).Value.ToString()
            cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(7).Value.ToString()
            cmd.Parameters.Add("@uom1", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(8).Value.ToString()
            cmd.Parameters.Add("@child", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(9).Value.ToString()
            cmd.Parameters.Add("@level", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(10).Value.ToString()
            cmd.Parameters.Add("@supply_type", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(11).Value.ToString()
            cmd.Parameters.Add("@description", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(12).Value.ToString()
            cmd.Parameters.Add("@planner", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(13).Value.ToString()
            cmd.Parameters.Add("@wh", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(14).Value.ToString()
            cmd.Parameters.Add("@sum_arrive", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(15).Value.ToString()
            cmd.Parameters.Add("@sum_departure", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(16).Value.ToString()
            cmd.Parameters.Add("@deliv_type", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(17).Value.ToString()
            cmd.Parameters.Add("@req_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(18).Value.ToString()
            cmd.Parameters.Add("@issued_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(19).Value.ToString()
            cmd.Parameters.Add("@remain_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(20).Value.ToString()
            cmd.Parameters.Add("@qpa", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(21).Value.ToString()
            cmd.Parameters.Add("@uit", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(22).Value.ToString()
            cmd.Parameters.Add("@uom2", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(23).Value.ToString()
            cmd.Parameters.Add("@raw_mtl", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(24).Value.ToString()
            cmd.Parameters.Add("@raw_wip", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(25).Value.ToString()
            cmd.Parameters.Add("@wip_assy", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(26).Value.ToString()
            cmd.Parameters.Add("@item_cost", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(27).Value.ToString()
            cmd.Parameters.Add("@line_code", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(28).Value.ToString()
            cmd.Parameters.Add("@substitute_item", MySqlDbType.VarChar).Value = DataGridView1.Rows(baris).Cells(29).Value.ToString()
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            cmd.ExecuteNonQuery()
        Next
        MsgBox("Data updated!", MsgBoxStyle.Information)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Label1.Text = DataGridView1.Rows(2).Cells(2).Value.ToString()
        Call SaveData()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim dt As DataTable = tables(ComboBox1.SelectedItem.ToString())
        DataGridView1.DataSource = dt
    End Sub
End Class