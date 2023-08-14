Imports MySql.Data.MySqlClient

Public Class FormLogLabelBoxCSKD

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Txn ID", "Txn ID") '0
        DataGridView1.Columns.Add("Type Box", "Type Box") '1
        DataGridView1.Columns.Add("WO", "WO") '2
        DataGridView1.Columns.Add("PN", "PN") '3
        DataGridView1.Columns.Add("Line", "Line") '4
        DataGridView1.Columns.Add("Shift", "Shift") '5
        DataGridView1.Columns.Add("PCBA ID", "PCBA ID") '6
        DataGridView1.Columns.Add("PIC", "PIC") '7
        DataGridView1.Columns.Add("Created", "Created") '8
        DataGridView1.Columns(0).Width = 150
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 100
        DataGridView1.Columns(7).Width = 120
        DataGridView1.Columns(8).Width = 790
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

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TXNID, type_box as TypeBox, qty_box as QtyBox, wo as WO, pn as PN, line as Line, shift as Shift, sn as PCBAID, pic as PIC, remark as Remark, created as Date FROM prodsys2_labelbox_cskd_log_tbl WHERE txn_id = '" & TextBox1.Text & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
    End Sub

    Private Sub FormLogLabelBoxCSKD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call showcount()
        Call customdgv2()
        Call BuatKolom()
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Try
                str = "SELECT txn_id as TXNID, type_box as TypeBox, qty_box as QtyBox, wo as WO, pn as PN, line as Line, shift as Shift, sn as PCBAID, pic as PIC, remark as Remark, created as Date FROM prodsys2_labelbox_cskd_log_tbl WHERE txn_id = '" & TextBox1.Text & "'"
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
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportDgvToExcel(Me.DataGridView1)
    End Sub
End Class