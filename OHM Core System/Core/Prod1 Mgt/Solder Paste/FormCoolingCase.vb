Imports MySql.Data.MySqlClient

Public Class FormCoolingCase

    Sub bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.SelectedItem = Nothing
        TextBox1.Select()
    End Sub

    Sub KolomCoolingCase()
        FormSolderPasteMGT.DataGridView1.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView1.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView1.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView1.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView1.Columns(4).Width = 180 'start time
        FormSolderPasteMGT.DataGridView1.Columns(5).Width = 180 'running time
        FormSolderPasteMGT.DataGridView1.Columns(6).Width = 180 'end time
        FormSolderPasteMGT.DataGridView1.Columns(7).Width = 120 'status
    End Sub

    Sub LoadDataCoolingCase()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCoolingCase()
    End Sub

    Sub ShowCountCoolingCase()
        FormSolderPasteMGT.Label3.Text = FormSolderPasteMGT.DataGridView1.RowCount
        FormSolderPasteMGT.TabPage1.Text = "Cooling Case [" & FormSolderPasteMGT.DataGridView1.RowCount & "]"
    End Sub

    Private Sub FormCoolingCase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call bersih()
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick
        TextBox2.SelectAll()
    End Sub

    Private Sub TextBox3_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox3.MouseClick
        TextBox3.SelectAll()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("DATA TIDAK LENGKAP!", vbInformation)
        Else
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_solder_paste_mgt_coolingcase_tbl` (`txn_id`, `maker`, `n_type`, `topbott` , `start_time`, `run_time`, `end_time`, `status`)
VALUES (@txn_id,@maker,@n_type,@topbott,@start_time, @run_time, @end_time, @status)"
            cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@maker", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@n_type", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
            cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
            cmd.Parameters.Add("@run_time", MySqlDbType.VarChar).Value = DBNull.Value
            cmd.Parameters.Add("@end_time", MySqlDbType.VarChar).Value = DBNull.Value
            cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
            cmd.CommandText = str
            Try
                conn.Open()
                cmd.ExecuteNonQuery()
                TextBox1.Text = ""
                TextBox1.Focus()
                Call LoadDataCoolingCase()
                Call ShowCountCoolingCase()
                rd.Close()
                conn.Close()
                conn.Dispose()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class