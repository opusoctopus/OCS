Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient

Public Class FormScreeningPCBAID
    Private Sub FormScreeningPCBAID_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call connection()
        Call bersih()
    End Sub

    Public Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Focus()
    End Sub

    Public Sub start()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox1.Focus()
    End Sub

    Public Sub StripNonAlphabetCharacters(ByVal input As TextBox)
        ' pattern matches any character that is NOT A-Z (allows upper and lower case alphabets)
        Dim rx As New Regex("[^0-9]")
        If (rx.IsMatch(input.Text)) Then
            Dim startPosition As Integer = input.SelectionStart - 1
            input.Text = rx.Replace(input.Text, "")
            input.SelectionStart = startPosition
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        StripNonAlphabetCharacters(TextBox1)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Len(TextBox2.Text) > 9 Then
                TextBox2.Text = TextBox2.Text.Substring(TextBox2.Text.Length - 8)
                Call connection()
                str = "SELECT DISTINCT ebt, wo, assembly_partno FROM prodsys_prod_plan_tbl WHERE wo='" & TextBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If Not rd.HasRows Then
                        Label7.Text = "Production plan hasn't been updated yet"
                        Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                        Label7.ForeColor = Color.Yellow
                        TextBox2.Text = ""
                        TextBox2.Focus()
                    Else
                        TextBox3.Text = rd("ebt")
                        TextBox4.Text = rd("assembly_partno")
                        TextBox5.Enabled = True
                        TextBox5.Focus()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try
            Else
                TextBox2.Text = ""
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            If TextBox5.Text = "Start" Then
                Call start()
                FormCountScreeningPCBAID.Show()
            ElseIf TextBox5.Text = "Stop" Then
                If Label9.Text = "0" Then
                    Label7.Text = "no SN is scanned"
                    Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                    Label7.ForeColor = Color.Yellow
                    Call bersih()
                ElseIf Label9.Text = Label11.Text Then
                    cmd = New MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_screening_pcbid_log_audit_tbl`(`txn_id`, `demand_id`, `qty_box`, `qty_actual`, `result`, `created`) VALUES (@txn_id,@demand_id,@qty_box,@qty_actual,@result,@created)"
                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                    cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                    cmd.Parameters.Add("@qty_box", MySqlDbType.VarChar).Value = Label11.Text
                    cmd.Parameters.Add("@qty_actual", MySqlDbType.VarChar).Value = Label9.Text
                    cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = "OK"
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        rd.Close()
                        cmd.Dispose()
                        conn.Dispose()
                        conn.Close()
                    End Try
                    Call bersih()
                Else
                    Label6.Text = "NG"
                    Label6.Font = New Font("Arial", 250, FontStyle.Bold)
                    Label6.ForeColor = Color.Red
                    Label7.Text = "Different Quantity"
                    Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                    Label7.ForeColor = Color.Yellow
                    My.Computer.Audio.Play("D:\ANDON-OHM\NG_OHM.wav")
                    cmd = New MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_screening_pcbid_log_audit_tbl`(`txn_id`, `demand_id`, `qty_box`, `qty_actual`, `result`, `created`) VALUES (@txn_id,@demand_id,@qty_box,@qty_actual,@result,@created)"
                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                    cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                    cmd.Parameters.Add("@qty_box", MySqlDbType.VarChar).Value = Label11.Text
                    cmd.Parameters.Add("@qty_actual", MySqlDbType.VarChar).Value = Label9.Text
                    cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = "NG"
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        rd.Close()
                        cmd.Dispose()
                        conn.Dispose()
                        conn.Close()
                    End Try
                    Call bersih()
                End If
            Else
                ' ======= CHECK DUPLICATE FAL ========
                Call connection()
                str = "SELECT pn, sn FROM prodsys2_sn_tbl WHERE sn='" & TextBox5.Text & "' UNION SELECT pn, sn FROM prodsys2_screening_pcbid_log_tbl WHERE sn='" & TextBox5.Text & "' ORDER BY pn"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        rd.Close()
                        Label6.Text = "NG"
                        Label6.Font = New Font("Arial", 250, FontStyle.Bold)
                        Label6.ForeColor = Color.Red
                        Label7.Text = "[" + TextBox5.Text + "]" + " DUPLICATE IN FAL"
                        Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                        Label7.ForeColor = Color.Yellow
                        My.Computer.Audio.Play("D:\ANDON-OHM\NG_OHM.wav")
                        ' ======= START SIMPAN LOG DUPLICATE FAL ========
                        cmd = New MySql.Data.MySqlClient.MySqlCommand
                        cmd.Connection = conn
                        str = "INSERT INTO `prodsys2_screening_pcbid_log_tbl`(`txn_id`, `model_suffix`, `pn`,`demand_id`, `sn`, `result_duplicate`, `remarks`, `inspector_lqc`,`created`) VALUES (@txn_id,@model_suffix,@pn,@demand_id,@sn,@result_duplicate,@remarks,@inspector_lqc,@created)"
                        cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                        cmd.Parameters.Add("@model_suffix", MySqlDbType.VarChar).Value = TextBox4.Text
                        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
                        cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                        cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox5.Text
                        cmd.Parameters.Add("@result_duplicate", MySqlDbType.VarChar).Value = Label7.Text
                        cmd.Parameters.Add("@remarks", MySqlDbType.VarChar).Value = Label6.Text
                        cmd.Parameters.Add("@inspector_lqc", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                        cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                        cmd.CommandText = str
                        Try
                            cmd.ExecuteNonQuery()
                            TextBox5.Text = ""
                            TextBox5.Focus()
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Finally
                            rd.Close()
                            cmd.Dispose()
                            conn.Close()
                            conn.Dispose()
                        End Try
                        ' ======= END SIMPAN LOG DUPLICATE FAL ========
                    Else
                        ' ======= CHECK NOT MIXING EBT ========
                        str = "SELECT pn, sn FROM prodsys_sn_tbl WHERE sn='" & TextBox5.Text & "'"
                        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                        rd = cmd.ExecuteReader
                        rd.Read()
                        Try
                            If rd.HasRows Then
                                rd.Close()
                                TextBox6.Text = rd("pn")
                                If TextBox3.Text = TextBox6.Text Then
                                    'KONDISI OK SEMUA
                                    Label6.Text = "OK"
                                    Label6.Font = New Font("Arial", 250, FontStyle.Bold)
                                    Label6.ForeColor = Color.Lime
                                    Label7.Text = ""
                                    Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                                    Label7.ForeColor = Color.Yellow
                                    My.Computer.Audio.Play("D:\ANDON-OHM\OK_OHM.wav")
                                    ' ======= START SIMPAN LOG OK ========
                                    cmd = New MySqlCommand
                                    cmd.Connection = conn
                                    str = "INSERT INTO `prodsys2_screening_pcbid_log_tbl`(`txn_id`, `model_suffix`, `pn`,`demand_id`, `sn`, `remarks`, `inspector_lqc`,`created`) VALUES (@txn_id,@model_suffix,@pn,@demand_id,@sn,@remarks,@inspector_lqc,@created)"
                                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                    cmd.Parameters.Add("@model_suffix", MySqlDbType.VarChar).Value = TextBox4.Text
                                    cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
                                    cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                                    cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox5.Text
                                    cmd.Parameters.Add("@remarks", MySqlDbType.VarChar).Value = Label6.Text
                                    cmd.Parameters.Add("@inspector_lqc", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                    cmd.CommandText = str
                                    Try
                                        cmd.ExecuteNonQuery()
                                        rd.Close()

                                        Label9.Text = "0"
                                        str = "select txn_id, count(sn) as qty from prodsys2_screening_pcbid_log_tbl where txn_id='" & TextBox1.Text & "'"
                                        cmd = New MySqlCommand(str, conn)
                                        rd = cmd.ExecuteReader
                                        rd.Read()
                                        If rd.HasRows Then
                                            Label9.Text = rd("qty")
                                        End If

                                        TextBox6.Text = ""
                                        TextBox5.Text = ""
                                        TextBox5.Focus()
                                    Catch ex As Exception
                                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Finally
                                        rd.Close()
                                    End Try
                                    ' ======= END SIMPAN LOG OK ========
                                Else
                                    Label6.Text = "NG"
                                    Label6.Font = New Font("Arial", 250, FontStyle.Bold)
                                    Label6.ForeColor = Color.Red
                                    Label7.Text = "[" + TextBox6.Text + "]" + " MIXING PARTNUMBER ASSY"
                                    Label7.Font = New Font("Arial", 24, FontStyle.Regular)
                                    Label7.ForeColor = Color.Yellow
                                    My.Computer.Audio.Play("D:\ANDON-OHM\NG_OHM.wav")
                                    ' ======= START SIMPAN LOG MIXING EBT ========
                                    cmd = New MySqlCommand
                                    cmd.Connection = conn
                                    str = "INSERT INTO `prodsys2_screening_pcbid_log_tbl`(`txn_id`, `model_suffix`, `pn`,`demand_id`, `sn`, `result_mixing`, `remarks`, `inspector_lqc`,`created`) VALUES (@txn_id,@model_suffix,@pn,@demand_id,@sn,@result_mixing,@remarks,@inspector_lqc,@created)"

                                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                    cmd.Parameters.Add("@model_suffix", MySqlDbType.VarChar).Value = TextBox4.Text
                                    cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox6.Text
                                    cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                                    cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox5.Text
                                    cmd.Parameters.Add("@result_mixing", MySqlDbType.VarChar).Value = Label7.Text
                                    cmd.Parameters.Add("@remarks", MySqlDbType.VarChar).Value = Label6.Text
                                    cmd.Parameters.Add("@inspector_lqc", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                    cmd.CommandText = str
                                    Try
                                        cmd.ExecuteNonQuery()
                                        TextBox5.Text = ""
                                        TextBox5.Focus()
                                    Catch ex As Exception
                                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Finally
                                        rd.Close()
                                    End Try
                                    ' ======= END SIMPAN LOG MIXING EBT ========
                                End If
                            Else
                                'KONDISI SN BELUM DI UPLOAD
                                Label6.Text = "NG"
                                Label6.Font = New Font("Arial", 250, FontStyle.Bold)
                                Label6.ForeColor = Color.Red
                                Label7.Text = "SN NOT REGISTERED"
                                Label7.Font = New Font("Arial", 72, FontStyle.Regular)
                                Label7.ForeColor = Color.Yellow
                                My.Computer.Audio.Play("D:\ANDON-OHM\NG_OHM.wav")
                                ' ======= START SIMPAN LOG SERIAL NUMBER BELUM DI UPLOAD ========
                                cmd = New MySqlCommand
                                cmd.Connection = conn
                                str = "INSERT INTO `prodsys2_screening_pcbid_log_tbl`(`txn_id`, `model_suffix`, `pn`,`demand_id`, `sn`, `result_mixing`, `remarks`, `inspector_lqc`,`created`) VALUES (@txn_id,@model_suffix,@pn,@demand_id,@sn,@result_mixing,@remarks,@inspector_lqc,@created)"
                                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                cmd.Parameters.Add("@model_suffix", MySqlDbType.VarChar).Value = TextBox4.Text
                                cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
                                cmd.Parameters.Add("@demand_id", MySqlDbType.VarChar).Value = TextBox2.Text
                                cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox5.Text
                                cmd.Parameters.Add("@result_mixing", MySqlDbType.VarChar).Value = Label7.Text
                                cmd.Parameters.Add("@remarks", MySqlDbType.VarChar).Value = Label6.Text
                                cmd.Parameters.Add("@inspector_lqc", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                cmd.CommandText = str
                                Try
                                    cmd.ExecuteNonQuery()
                                    TextBox6.Text = ""
                                    TextBox5.Text = ""
                                    TextBox5.Focus()
                                Catch ex As Exception
                                    MessageBox.Show("ERROR" & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Finally
                                    rd.Close()
                                End Try
                                ' ======= END SIMPAN LOG SERIAL NUMBER BELUM DI UPLOAD ========
                            End If
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try
                TextBox5.Text = ""
                TextBox5.Focus()
            End If
        End If
    End Sub
End Class