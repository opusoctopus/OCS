Public Class FormListICeMMC
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
        DataGridView1.Columns(0).Width = 150
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 143
    End Sub

    Private Sub FormListICeMMC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT pn_comp, version, checksum, marking, qty, worker, created FROM prodsys2_ic_mapping_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        Call buatkolom()
        TextBox1.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormRequestICeMMC.TextBox8.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' PN COMP
        FormRequestICeMMC.TextBox7.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' VERSION
        FormRequestICeMMC.TextBox6.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' CHECKSUM
        FormRequestICeMMC.TextBox5.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' MARKING
        FormRequestICeMMC.TextBox9.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' MARKING

        Call connection()
        str = "SELECT * FROM prodsys2_ic_mapping_tbl WHERE pn_comp = '" & DataGridView1.SelectedCells(0).Value.ToString & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                FormRequestICeMMC.Label12.Text = rd("id")
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
        conn.Dispose()
        Me.Close()
    End Sub
End Class