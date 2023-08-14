Imports MySql.Data.MySqlClient

Public Class FormFeederData

    Dim cmd As MySqlCommand
    Dim dr As MySqlDataReader
    Dim sql As String
    Dim result As Integer

    Sub Bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.SelectedItem = Nothing
        TextBox1.Select()
    End Sub

    Private Sub FormFeederData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Bersih()

        Try
            conn.Open()
            sql = "SELECT sn, machine FROM prodsys_feeder_tbl"
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                dr = .ExecuteReader()
            End With
            ListView1.Items.Clear()
            Do While dr.Read = True
                Dim list = ListView1.Items.Add(dr(0))
                list.SubItems.Add(dr(1))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_feeder_tbl`(`sn`, `type`, `machine` , `created`) VALUES (@sn,@type,@machine,@created)"
        cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@machine", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@type", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
            rd.Close()
            conn.Close()
            conn.Dispose()
            Call Bersih()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class