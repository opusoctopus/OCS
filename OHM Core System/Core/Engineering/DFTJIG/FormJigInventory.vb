Imports MySql.Data.MySqlClient

Public Class FormJigInventory

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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(200, 5, 83)
            .DefaultCellStyle.Font = New System.Drawing.Font("Arial", 10.0!)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 24
            End With
        End With
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 60 '
        DataGridView1.Columns(1).Width = 140 '
        DataGridView1.Columns(2).Width = 120 '
        DataGridView1.Columns(3).Width = 120 '
        DataGridView1.Columns(4).Width = 140 '
        DataGridView1.Columns(5).Width = 140 '
        DataGridView1.Columns(6).Width = 100 '
        DataGridView1.Columns(7).Width = 100 '
        DataGridView1.Columns(8).Width = 180 '
        DataGridView1.Columns(9).Width = 180 '
        DataGridView1.Columns(10).Width = 180 '
        DataGridView1.Columns(11).Width = 180 '
        DataGridView1.Columns(12).Width = 180 '
    End Sub

    Sub ShowCount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Sub Bersih()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()

        ComboBox1.SelectedItem = Nothing

    End Sub

    Sub LoadData()
        Try
            str = "SELECT chasis, kode_jig, maker, kode_rak, checkpoint1, checkpoint2, checkpoint3, checkpoint4, checkpoint5, doc_bc, DATE_FORMAT(STR_TO_DATE(incoming_date,'%d/%m/%Y'), '%Y/%m/%d') as incoming_date, DATE_FORMAT(STR_TO_DATE(return_date,'%d/%m/%Y'), '%Y/%m/%d') as return_date, created FROM prodsys2_jig_inventory_tbl ORDER BY incoming_date"
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

    Private Sub FormJigInventory_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call customdgv2()
        Call LoadData()
    End Sub
End Class