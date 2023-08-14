Imports MySql.Data.MySqlClient

Public Class FormFeederInspection

    Sub bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        ComboBox3.SelectedItem = Nothing
        ComboBox4.SelectedItem = Nothing
        ComboBox5.SelectedItem = Nothing
        ComboBox6.SelectedItem = Nothing
        ComboBox7.SelectedItem = Nothing
        ComboBox8.SelectedItem = Nothing
        ComboBox9.SelectedItem = Nothing
        ComboBox10.SelectedItem = Nothing
        ComboBox11.SelectedItem = Nothing
        ComboBox12.SelectedItem = Nothing
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.SelectedItem = "CALIBRATION" Then
            conn.Open()
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_feeder_log_tbl`(`sn`,`machine`, `category`, `cp1` , `cp2`, `cp3`, `cp4`, `cp5`, `cp6`, `cp7`, `cp8`, `cp9`, `cp10`, `throuble`, `status`, `pic`, `next_calibration`, `created`) VALUES (@sn,@machine,@category,@cp1,@cp2,@cp3,@cp4,@cp5,@cp6,@cp7,@cp8,@cp9,@cp10,@throuble,@status,@pic,@next_calibration,@created)"
            cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@machine", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@category", MySqlDbType.VarChar).Value = ComboBox1.Text
            cmd.Parameters.Add("@cp1", MySqlDbType.VarChar).Value = ComboBox2.Text
            cmd.Parameters.Add("@cp2", MySqlDbType.VarChar).Value = ComboBox3.Text
            cmd.Parameters.Add("@cp3", MySqlDbType.VarChar).Value = ComboBox4.Text
            cmd.Parameters.Add("@cp4", MySqlDbType.VarChar).Value = ComboBox5.Text
            cmd.Parameters.Add("@cp5", MySqlDbType.VarChar).Value = ComboBox6.Text
            cmd.Parameters.Add("@cp6", MySqlDbType.VarChar).Value = ComboBox11.Text
            cmd.Parameters.Add("@cp7", MySqlDbType.VarChar).Value = ComboBox10.Text
            cmd.Parameters.Add("@cp8", MySqlDbType.VarChar).Value = ComboBox9.Text
            cmd.Parameters.Add("@cp9", MySqlDbType.VarChar).Value = ComboBox8.Text
            cmd.Parameters.Add("@cp10", MySqlDbType.VarChar).Value = ComboBox7.Text
            cmd.Parameters.Add("@throuble", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = ComboBox12.Text
            cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
            cmd.Parameters.Add("@next_calibration", MySqlDbType.VarChar).Value = DateTimePicker1.Text
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
            cmd.CommandText = str
            Try
                cmd.ExecuteNonQuery()
                conn.Close()
                Call bersih()
                TextBox1.Select()
                rd.Close()
                conn.Close()
                conn.Dispose()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        ElseIf ComboBox1.SelectedItem = "REPAIR" Then
            conn.Open()
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_feeder_log_tbl`(`sn`,`machine`, `category`, `cp1` , `cp2`, `cp3`, `cp4`, `cp5`, `cp6`, `cp7`, `cp8`, `cp9`, `cp10`, `throuble`, `status`, `pic`, `next_calibration`, `created`) VALUES (@sn,@machine,@category,@cp1,@cp2,@cp3,@cp4,@cp5,@cp6,@cp7,@cp8,@cp9,@cp10,@throuble,@status,@pic,@next_calibration,@created)"
            cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@machine", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@category", MySqlDbType.VarChar).Value = ComboBox1.Text
            cmd.Parameters.Add("@cp1", MySqlDbType.VarChar).Value = ComboBox2.Text
            cmd.Parameters.Add("@cp2", MySqlDbType.VarChar).Value = ComboBox3.Text
            cmd.Parameters.Add("@cp3", MySqlDbType.VarChar).Value = ComboBox4.Text
            cmd.Parameters.Add("@cp4", MySqlDbType.VarChar).Value = ComboBox5.Text
            cmd.Parameters.Add("@cp5", MySqlDbType.VarChar).Value = ComboBox6.Text
            cmd.Parameters.Add("@cp6", MySqlDbType.VarChar).Value = ComboBox11.Text
            cmd.Parameters.Add("@cp7", MySqlDbType.VarChar).Value = ComboBox10.Text
            cmd.Parameters.Add("@cp8", MySqlDbType.VarChar).Value = ComboBox9.Text
            cmd.Parameters.Add("@cp9", MySqlDbType.VarChar).Value = ComboBox8.Text
            cmd.Parameters.Add("@cp10", MySqlDbType.VarChar).Value = ComboBox7.Text
            cmd.Parameters.Add("@throuble", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = ComboBox12.Text
            cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
            cmd.Parameters.Add("@next_calibration", MySqlDbType.VarChar).Value = DBNull.Value
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
            cmd.CommandText = str
            Try
                cmd.ExecuteNonQuery()
                conn.Close()
                Call bersih()
                TextBox1.Select()
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

    Private Sub FormFeederInspection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        Label16.Hide()
        DateTimePicker1.Hide()

        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(365)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedItem = "CALIBRATION" Then
            Label16.Show()
            DateTimePicker1.Show()
        Else
            Label16.Hide()
            DateTimePicker1.Hide()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT * FROM prodsys2_feeder_tbl WHERE sn='" & TextBox1.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If Not rd.HasRows Then
                    MsgBox("Feeder ini belum terdaftar!", vbExclamation)
                    TextBox1.Clear()
                    TextBox1.Select()
                    Call bersih()
                Else
                    'TextBox1.Text = rd("txn_id")
                    TextBox2.Text = rd("machine")
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            Finally
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub
End Class