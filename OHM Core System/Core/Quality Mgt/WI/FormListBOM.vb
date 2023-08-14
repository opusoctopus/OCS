Public Class FormListBOM
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
            .DefaultCellStyle.SelectionBackColor = Color.DarkCyan
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.DimGray
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Sub buatkolom()
        DataGridView1.Columns(0).Width = 335
    End Sub

    Sub LoadData()
        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT DISTINCT id_bom FROM prodsys2_master_bom_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        Call buatkolom()
        TextBox1.Focus()
    End Sub
    Sub showcount()
        Label3.Text = DataGridView1.RowCount
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FormListBOM_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call customdgv2()
        Call LoadData()
        Call showcount()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            rd.Close()
            da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT DISTINCT id_bom FROM prodsys2_master_bom_tbl WHERE id_bom = '" & TextBox1.Text & "'", conn)
            ds = New DataSet
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Call showcount()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call LoadData()
        Call showcount()
        TextBox1.Text = ""
        TextBox1.Focus()
    End Sub
End Class