Imports MySql.Data.MySqlClient
Public Class FormRepairCodeMaster

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
        DataGridView1.Columns(0).Width = 80 'STATUS
        DataGridView1.Columns(1).Width = 80 'CODE
        DataGridView1.Columns(2).Width = 120 'DESC
        DataGridView1.Columns(3).Width = 100 'NAME
        DataGridView1.Columns(4).Width = 100 'CREATED
        DataGridView1.Columns(5).Width = 120 'UPDATED
        DataGridView1.Columns(6).Width = 120 'UPDATED
    End Sub

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT * FROM prodsys2_repair_code_master_tbl"
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

    Private Sub FormRepairCodeMaster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call showcount()
        Call BuatKolom()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormRepairInputCodeMaster.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportDgvToExcel(Me.DataGridView1)
    End Sub
End Class