Imports MySql.Data.MySqlClient

Public Class FormRepairInputCodeMaster
    Private Sub FormRepairInputCodeMaster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Text = FormMain.LabelNama.Text
        TextBox3.ReadOnly = True
        CheckBox1.Checked = True
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Sub BuatKolom()
        FormRepairCodeMaster.DataGridView1.Columns(0).Width = 80 'STATUS
        FormRepairCodeMaster.DataGridView1.Columns(1).Width = 80 'CODE
        FormRepairCodeMaster.DataGridView1.Columns(2).Width = 120 'DESC
        FormRepairCodeMaster.DataGridView1.Columns(3).Width = 100 'NAME
        FormRepairCodeMaster.DataGridView1.Columns(4).Width = 100 'CREATED
        FormRepairCodeMaster.DataGridView1.Columns(5).Width = 120 'UPDATED
        FormRepairCodeMaster.DataGridView1.Columns(6).Width = 120 'UPDATED
    End Sub

    Sub showcount()
        FormRepairCodeMaster.Label22.Text = FormRepairCodeMaster.DataGridView1.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT * FROM prodsys2_repair_code_master_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormRepairCodeMaster.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_repair_code_master_tbl` (`v`, `code`, `desc_code`, `c_name` , `c_date`)
VALUES (@v,@code,@desc_code,@c_name,@c_date)"

        cmd.Parameters.Add("@v", MySqlDbType.VarChar).Value = "Y"
        cmd.Parameters.Add("@code", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@desc_code", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@c_name", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@c_date", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            Call BuatKolom()
            Call showcount()
            Call LoadData()
            rd.Close()
            conn.Close()
            conn.Dispose()
            TextBox2.Text = ""
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub
End Class