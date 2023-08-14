Public Class FormUploadData
    Dim eventhandleradded As Boolean = False

    Sub customdgv2()
        With DataGridView1 'Ganti dengan nama datagridview kalian
            '.AllowUserToAddRows = False
            .RowHeadersVisible = False
            .BorderStyle = BorderStyle.None
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 30
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

        With DataGridView2 'Ganti dengan nama datagridview kalian
            '.AllowUserToAddRows = False
            .RowHeadersVisible = False
            .BorderStyle = BorderStyle.None
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 30
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

    Private Sub DataGridView2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseClick
        'if the datagridview was right clicked, set the currentcell to the right-clicked cell and then show a context menu with options for filling the right-clicked cell
        If (e.Button = MouseButtons.Right) Then
            Try
                'highlight the right-clicked cell, make it the dgv's current cell
                DataGridView2.CurrentRow.Selected = False
                DataGridView2.Rows(e.RowIndex).Selected = True
                DataGridView2.CurrentCell = DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex)
                DataGridView2.ContextMenuStrip = ContextMenuStrip1
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub FormUploadData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()

        DataGridView2.ClearSelection()
        DataGridView2.CurrentCell = DataGridView2.Item("Column1", 0)
        DataGridView2.BeginEdit(True)
        Dim row As String() = New String() {" ", " ", " "}
        DataGridView2.Rows.Add(row)
        DataGridView2.ClearSelection()
        DataGridView2.Rows(0).Cells(0).Selected = True
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        CopyToClipboard()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        'Copy to clipboard
        CopyToClipboard()
        'Clear selected cells
        For Each dgvCell As DataGridViewCell In DataGridView2.SelectedCells
            dgvCell.Value = String.Empty
        Next
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        ' Perform paste Operation
        PasteClipboardValue()
    End Sub

    Private Sub CopyToClipboard()
        ' Copy to clipboard
        Dim dataObj As DataObject = DataGridView2.GetClipboardContent()
        If (Not IsNothing(dataObj)) Then
            Clipboard.SetDataObject(dataObj)
        End If
    End Sub

    Private Sub PasteClipboardValue()
        ' Show Error if no cell is selected
        If (DataGridView2.SelectedCells.Count = 0) Then
            MessageBox.Show("Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get the starting Cell
        Dim startCell As DataGridViewCell = GetStartCell(DataGridView2)
        ' Get the clipboard value in a dictionary
        Dim cbValue As New Dictionary(Of Integer, Dictionary(Of Integer, String))
        cbValue = ClipBoardValues(Clipboard.GetText())

        Dim iRowIndex As Integer = startCell.RowIndex
        For Each rowKey As Integer In cbValue.Keys
            Dim iColIndex As Integer = startCell.ColumnIndex
            For Each cellKey As Integer In cbValue(rowKey).Keys
                ' Check if the index is within the limit
                If (iColIndex <= DataGridView2.Columns.Count - 1 And
                iRowIndex <= DataGridView2.Rows.Count - 1) Then
                    Dim item() As Object = cbValue(rowKey)(cellKey).Split(vbTab)
                    For j As Integer = 0 To item.Count - 1
                        Dim cell As DataGridViewCell = DataGridView2(iColIndex + j, iRowIndex)
                        cell.Value = item(j)

                    Next
                End If
                iColIndex += 1

            Next
            iRowIndex += 1
            DataGridView2.Rows.Add()
        Next

        iRowIndex = 0

    End Sub
    Private Function GetStartCell(dgView As DataGridView) As DataGridViewCell
        ' get the smallest row,column index
        If (dgView.SelectedCells.Count = 0) Then
            Return Nothing
        End If
        Dim rowIndex As Integer = dgView.Rows.Count - 1
        Dim colIndex As Integer = dgView.Columns.Count - 1

        For Each dgvCell As DataGridViewCell In dgView.SelectedCells
            If (dgvCell.RowIndex < rowIndex) Then
                rowIndex = dgvCell.RowIndex
            End If
            If (dgvCell.ColumnIndex < colIndex) Then
                colIndex = dgvCell.ColumnIndex
            End If
        Next
        Return dgView(colIndex, rowIndex)
    End Function

    Private Function ClipBoardValues(clipboardValue As String) As Dictionary(Of Integer, Dictionary(Of Integer, String))
        Dim copyValues As Dictionary(Of Integer, Dictionary(Of Integer, String)) =
        New Dictionary(Of Integer, Dictionary(Of Integer, String))()

        Dim lines As String() = clipboardValue.Split(vbCrLf)

        For i As Integer = 0 To lines.Length - 1
            copyValues(i) = New Dictionary(Of Integer, String)()
            Dim lineContent As String() = lines(i).Split("\t")

            ' if an empty cell value copied, then set the dictionary with an empty string
            ' else Set value to dictionary
            If (lineContent.Length = 0) Then
                copyValues(i)(0) = String.Empty
            Else
                For j As Integer = 0 To lineContent.Length - 1
                    copyValues(i)(j) = lineContent(j)
                Next
            End If
        Next
        Return copyValues
    End Function

    Private Sub DataGridView2_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing
        If eventhandleradded = False Then
            AddHandler e.Control.KeyDown, AddressOf cell_Keydown
            eventhandleradded = True
        End If
    End Sub

    Private Sub cell_Keydown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.P Then
            e.Handled = True
            e.SuppressKeyPress = True
            PasteClipboardValue()
        ElseIf e.KeyCode = Keys.X Then
            e.Handled = True
            e.SuppressKeyPress = True
            CopyToClipboard()
            For Each dgvCell As DataGridViewCell In DataGridView2.SelectedCells
                dgvCell.Value = String.Empty
            Next
        ElseIf e.KeyCode = Keys.C Then
            e.Handled = True
            e.SuppressKeyPress = True
            CopyToClipboard()
        End If
    End Sub

    Private Sub UploadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UploadToolStripMenuItem.Click

    End Sub
End Class

