Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class FormNozzleMGT

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
        With DataGridView2 'Ganti dengan nama datagridview kalian
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

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID,
                    nozzle_type as NozzleType, 
                    status_inspection as Judgment,
                    remark as Remark,
                    pic as PIC,
                    approved as Approved,
                    approval as Signed,
                    DATE_FORMAT(STR_TO_DATE(created,'%d/%m/%Y'), '%Y/%m/%d') as Created
                   FROM prodsys2_nozzle_inspection_tbl ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolom()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT line as Line,
                    machine as Machine,   
                    nozzle_type as NozzleType,
                    count(nozzle_type) as QTY
                   FROM prodsys2_nozzle_mgt_log_tbl ORDER BY line ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInLine()
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 110
        DataGridView1.Columns(2).Width = 90
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(4).Width = 120
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).Width = 180
    End Sub

    Sub BuatKolomInLine()
        DataGridView2.Columns(0).Width = 120
        DataGridView2.Columns(1).Width = 130
        DataGridView2.Columns(2).Width = 120
        DataGridView2.Columns(3).Width = 380
    End Sub

    Sub ShowCout()
        Label2.Text = DataGridView1.RowCount
        Label3.Text = DataGridView2.RowCount
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormNozzleInventory.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormNozzleInspection.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormNozzleOut.ShowDialog()
    End Sub

    Private Sub FormNozzleMGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call LoadDataInLine()
        Call ShowCout()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormNozzleDetailInspection.Label6.Text = DataGridView1.SelectedCells(0).Value.ToString
        FormNozzleDetailInspection.TextBox1.Text = DataGridView1.SelectedCells(1).Value.ToString
        FormNozzleDetailInspection.TextBox2.Text = DataGridView1.SelectedCells(2).Value.ToString
        FormNozzleDetailInspection.TextBox3.Text = DataGridView1.SelectedCells(3).Value.ToString
        FormNozzleDetailInspection.TextBox4.Text = DataGridView1.SelectedCells(4).Value.ToString

        FormNozzleDetailInspection.ShowDialog()
    End Sub
End Class