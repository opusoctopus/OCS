Imports MySql.Data.MySqlClient

Public Class FormMSL3Data

    Dim cmd As MySqlCommand
    Dim dr As MySqlDataReader
    Dim sql As String
    Dim result As Integer

    Sub Bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()

        TextBox2.Select()
    End Sub

    Sub LoadData()
        Call Bersih()

        Try
            Call connection()
            sql = "SELECT unique_id, pn_lg, lot_id, qty, incoming_date FROM prodsys2_master_msl3_tbl"
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
                list.SubItems.Add(dr(2))
                list.SubItems.Add(dr(3))
                list.SubItems.Add(dr(4))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Sub NomorOtomatis()
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM prodsys2_master_msl3_tbl WHERE unique_id IN (SELECT MAX(unique_id) FROM prodsys2_master_msl3_tbl)", conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            UrutanKode = "MSL3" + System.DateTime.Now.ToString("yyyyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            UrutanKode = "MSL3" + System.DateTime.Now.ToString("yyyyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = UrutanKode
        rd.Close()
        conn.Close()
    End Sub

    Sub showcount()
        Label7.Text = ListView1.Items.Count
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Data tidak lengkap!", vbExclamation)
            TextBox2.Select()
        Else
            conn.Open()
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_master_msl3_tbl`(`unique_id`, `pn_lg`, `lot_id`, `qty`, `incoming_date`) VALUES (@unique_id,@pn_lg,@lot_id,@qty,@incoming_date)"
            cmd.Parameters.Add("@unique_id", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@pn_lg", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@lot_id", MySqlDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox5.Text
            cmd.Parameters.Add("@incoming_date", MySqlDbType.VarChar).Value = DateTimePicker1.Text
            cmd.CommandText = str
            Try
                cmd.ExecuteNonQuery()
                rd.Close()
                conn.Close()
                conn.Dispose()
                Call LoadData()
                Call NomorOtomatis()
                Call showcount()
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

    Private Sub FormMSL3Data_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        Call LoadData()
        Call NomorOtomatis()
        Call showcount()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox5.Focus()
        End If
    End Sub
End Class