Imports MySql.Data.MySqlClient

Public Class FormReturnMaterialNG
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub

    Sub LoadData()
        Try
            str = "SELECT wo, pn, pn_comp, qty, requested, remark, created FROM prodsys2_return_material_ng_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom2()
        conn.Close()
    End Sub

    Sub showcount()
        Label8.Text = DataGridView2.RowCount
        Label10.Text = DataGridView1.RowCount
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 50
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(3).Width = 180
        DataGridView1.Columns(4).Width = 200
        DataGridView1.Columns(5).Width = 150
    End Sub

    Sub BuatKolom2()
        DataGridView2.Columns(0).Width = 100
        DataGridView2.Columns(1).Width = 100
        DataGridView2.Columns(2).Width = 100
        DataGridView2.Columns(3).Width = 60
        DataGridView2.Columns(4).Width = 120
        DataGridView2.Columns(5).Width = 120
        DataGridView2.Columns(6).Width = 120
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
            .RowsDefaultCellStyle.SelectionForeColor = Color.Black
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
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
            .RowsDefaultCellStyle.SelectionForeColor = Color.Black
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            str = "SELECT distinct level, child, description, specification, designators FROM prodsys2_master_bom_tbl WHERE child = '" & TextBox1.Text & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        conn.Close()
    End Sub

    Private Sub FormReturnMaterialNG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call bersih()
        Call LoadData()
        Call showcount()

        Dim checkboxcolumn As New DataGridViewCheckBoxColumn
        checkboxcolumn.Width = 70
        checkboxcolumn.Name = "checkboxcolumn"
        checkboxcolumn.HeaderText = "Select"
        DataGridView1.Columns.Insert(0, checkboxcolumn)
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim table2 As New DataTable
        With table2
            .Columns.Add("Level")
            .Columns.Add("Part Number")
            .Columns.Add("Part Name")
            .Columns.Add("Specification")
            .Columns.Add("Location")
        End With

        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim isselect As Boolean = Convert.ToBoolean(row.Cells("checkboxcolumn").Value)
            If isselect Then
                TextBox4.Text = row.Cells(2).Value
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormProdPlanReturnMaterialNG.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_return_material_ng_tbl`(`wo`, `pn`, `pn_comp`, `qty`, `requested`, `remark`, `created`) VALUES (@wo,@pn,@pn_comp,@qty,@requested,@remark,@created)"
        cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@pn_comp", MySqlDbType.VarChar).Value = TextBox4.Text
        cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox5.Text
        cmd.Parameters.Add("@requested", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
        cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox6.Text
        cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call bersih()
        Call LoadData()
        Call showcount()
        conn.Close()
    End Sub
End Class