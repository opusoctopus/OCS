Imports MySql.Data.MySqlClient
Public Class FormMasterRepair
    Sub customdgv2()
        With DataGridView1 'Ganti dengan nama datagridview kalian
            .AllowUserToAddRows = False
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

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 120 'TXN ID
        DataGridView1.Columns(1).Width = 120 'STATUS
        DataGridView1.Columns(2).Width = 120 'ORG
        DataGridView1.Columns(3).Width = 120 'WO
        DataGridView1.Columns(4).Width = 120 'PN
        DataGridView1.Columns(5).Width = 180 'MODEL
        DataGridView1.Columns(6).Width = 80 'LINE
        DataGridView1.Columns(7).Width = 80 'TEAM
        DataGridView1.Columns(8).Width = 60 'QTY
        DataGridView1.Columns(9).Width = 80 'PROD
        DataGridView1.Columns(10).Width = 100 'SN
        DataGridView1.Columns(11).Width = 180 'CAUSE
        DataGridView1.Columns(12).Width = 180 'CAUSE2
        DataGridView1.Columns(13).Width = 120 'REQUESTOR
        DataGridView1.Columns(14).Width = 120 'REPAIRMAN
        DataGridView1.Columns(15).Width = 120 'POSITION
        DataGridView1.Columns(16).Width = 120 'COMPONENT
        DataGridView1.Columns(17).Width = 120 'REMARK
        DataGridView1.Columns(18).Width = 120 'SMT REPAIR
        DataGridView1.Columns(19).Width = 120 'DRM REPAIR
        DataGridView1.Columns(20).Width = 120 'DFT REPAIR
        DataGridView1.Columns(21).Width = 120 'VERIFIED STATUS
        DataGridView1.Columns(22).Width = 130 'CREATED
        DataGridView1.Columns(23).Width = 130 'UPDATED
    End Sub

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TXNID, status as Status, org as ORG, wo as WO, pn as PartNumber, model as Model, line as Line, team as Team, qty as Qty, product as Product, sn as BarCode, 
cause as Cause, cause2 as Cause2, requestor as Requestor, repairman as Repairman, position as Position, component as Component, remark as Remark, smt_repair as SMT, drm_repair as DRM, dft_repair as DFT, verified_status as Verified, created as Created, updated as Updated
FROM prodsys2_repair_management_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()
    End Sub

    Sub LoadDataPencarian()
        Try
            str = "SELECT txn_id as TXNID, status as Status, org as ORG, wo as WO, pn as PartNumber, model as Model, line as Line, team as Team, qty as Qty, product as Product, sn as BarCode, 
cause as Cause, cause2 as Cause2, requestor as Requestor, repairman as Repairman, position as Position, component as Component, remark as Remark, smt_repair as SMT, drm_repair as DRM, dft_repair as DFT, verified_status as Verified, created as Created, updated as Updated
FROM prodsys2_repair_management_tbl WHERE sn LIKE '" & TextBox1.Text & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormRepairInput.ShowDialog()
        FormRepairInput.TextBox5.Focus()
    End Sub

    Private Sub FormMasterRepair_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call customdgv2()
        Call LoadData()
        Call showcount()
        Call BuatKolom()
        ComboBox5.Text = "PROGRESS"

        Dim status As Boolean = False
        For Each row As DataGridViewRow In DataGridView1.Rows
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                If row.Cells(i).Value.ToString = ComboBox5.Text Then
                    status = True
                    row.Cells(i).Style.ForeColor = Color.Red
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormRepairUpdate.TextBox11.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' TXN ID
        FormRepairUpdate.ComboBox1.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' STATUS
        FormRepairUpdate.ComboBox2.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' ORG
        FormRepairUpdate.TextBox1.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' WO
        FormRepairUpdate.TextBox2.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' PN
        FormRepairUpdate.TextBox3.Text = Me.DataGridView1.SelectedCells(5).Value.ToString ' MODEL
        FormRepairUpdate.ComboBox3.Text = Me.DataGridView1.SelectedCells(6).Value.ToString ' LINE
        FormRepairUpdate.ComboBox4.Text = Me.DataGridView1.SelectedCells(7).Value.ToString ' SHIFT
        FormRepairUpdate.TextBox4.Text = Me.DataGridView1.SelectedCells(8).Value.ToString ' QTY

        If Me.DataGridView1.SelectedCells(9).Value.ToString = "SMT" Then ' PRODUCT
            FormRepairUpdate.CheckBox1.Checked = True ' SMT
        ElseIf Me.DataGridView1.SelectedCells(9).Value.ToString = "DRM" Then
            FormRepairUpdate.CheckBox2.Checked = True ' SMT
        ElseIf Me.DataGridView1.SelectedCells(9).Value.ToString = "DFT" Then
            FormRepairUpdate.CheckBox3.Checked = True ' SMT
        End If


        FormRepairUpdate.TextBox5.Text = Me.DataGridView1.SelectedCells(10).Value.ToString ' SN
        FormRepairUpdate.TextBox13.Text = Me.DataGridView1.SelectedCells(11).Value.ToString ' CAUSE
        FormRepairUpdate.ComboBox5.Text = Me.DataGridView1.SelectedCells(12).Value.ToString ' CAUSE2
        FormRepairUpdate.TextBox6.Text = Me.DataGridView1.SelectedCells(13).Value.ToString ' REQUESTOR
        FormRepairUpdate.TextBox10.Text = Me.DataGridView1.SelectedCells(14).Value.ToString ' REPAIRMAN
        FormRepairUpdate.TextBox7.Text = Me.DataGridView1.SelectedCells(15).Value.ToString ' POSITION
        FormRepairUpdate.TextBox8.Text = Me.DataGridView1.SelectedCells(16).Value.ToString ' COMPONEN
        FormRepairUpdate.TextBox9.Text = Me.DataGridView1.SelectedCells(17).Value.ToString ' REMARK
        If Me.DataGridView1.SelectedCells(21).Value.ToString = "1" Then
            FormRepairUpdate.CheckBox4.Checked = True
            FormRepairUpdate.CheckBox4.Enabled = False
        Else
            FormRepairUpdate.CheckBox4.Checked = False
            FormRepairUpdate.CheckBox4.Enabled = True
        End If
        FormRepairUpdate.ShowDialog()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox1.Text.StartsWith("IO")) Then
                MsgBox("SN [" & TextBox1.Text & "] INVALID CHARACTER", MsgBoxStyle.Information)
                Call customdgv2()
                Call LoadData()
                Call showcount()
                Call BuatKolom()
                TextBox1.SelectAll()
            Else
                Call customdgv2()
                Call LoadDataPencarian()
                Call showcount()
                Call BuatKolom()
                TextBox1.SelectAll()
            End If
        End If
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.SelectAll()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportDgvToExcel(Me.DataGridView1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call customdgv2()
        Call LoadData()
        Call showcount()
        Call BuatKolom()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call LoadDataPencarian()
    End Sub
End Class