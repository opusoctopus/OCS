﻿Imports MySql.Data.MySqlClient

Public Class FormLine

    Sub bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
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

    Sub KolomCheckRight()
        FormSolderPasteMGT.DataGridView2.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView2.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView2.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView2.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView2.Columns(4).Width = 120 'hole
        FormSolderPasteMGT.DataGridView2.Columns(5).Width = 180 'start time
        FormSolderPasteMGT.DataGridView2.Columns(6).Width = 180 'running time
        FormSolderPasteMGT.DataGridView2.Columns(7).Width = 180 'end time
        FormSolderPasteMGT.DataGridView2.Columns(8).Width = 180 'status
        FormSolderPasteMGT.DataGridView2.Columns(9).Width = 180 'pic
    End Sub

    Sub KolomMixer()
        FormSolderPasteMGT.DataGridView3.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView3.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView3.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView3.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView3.Columns(4).Width = 180 'start time
        FormSolderPasteMGT.DataGridView3.Columns(5).Width = 180 'running time
        FormSolderPasteMGT.DataGridView3.Columns(6).Width = 180 'end time
        FormSolderPasteMGT.DataGridView3.Columns(7).Width = 180 'status
    End Sub

    Sub KolomLine()
        FormSolderPasteMGT.DataGridView4.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView4.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView4.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView4.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView4.Columns(4).Width = 120 'line
        FormSolderPasteMGT.DataGridView4.Columns(5).Width = 180 'start time
        FormSolderPasteMGT.DataGridView4.Columns(6).Width = 180 'running time
        FormSolderPasteMGT.DataGridView4.Columns(7).Width = 180 'end time
        FormSolderPasteMGT.DataGridView2.Columns(8).Width = 180 'status
        FormSolderPasteMGT.DataGridView4.Columns(9).Width = 180 'pic
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

    Sub LoadDataCheckRight()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, hole, start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCheckRight()
    End Sub

    Sub LoadDataMixer()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomMixer()
    End Sub

    Sub LoadDataLine()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, line, start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView4.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLine()
    End Sub

    Sub ShowCountCoolingCase()
        FormSolderPasteMGT.Label3.Text = FormSolderPasteMGT.DataGridView1.RowCount
        FormSolderPasteMGT.TabPage1.Text = "Cooling Case [" & FormSolderPasteMGT.DataGridView1.RowCount & "]"
    End Sub

    Sub ShowCountCheckRight()
        FormSolderPasteMGT.Label6.Text = FormSolderPasteMGT.DataGridView2.RowCount
        FormSolderPasteMGT.TabPage2.Text = "Check Right [" & FormSolderPasteMGT.DataGridView2.RowCount & "]"
    End Sub

    Sub ShowCountMixer()
        FormSolderPasteMGT.Label11.Text = FormSolderPasteMGT.DataGridView3.RowCount
        FormSolderPasteMGT.TabPage3.Text = "Mixer [" & FormSolderPasteMGT.DataGridView3.RowCount & "]"
    End Sub

    Sub ShowCountLine()
        FormSolderPasteMGT.Label7.Text = FormSolderPasteMGT.DataGridView4.RowCount
        FormSolderPasteMGT.TabPage4.Text = "Line [" & FormSolderPasteMGT.DataGridView4.RowCount & "]"
    End Sub

    Private Sub FormLine_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call bersih()
        TextBox1.Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Len(TextBox1.Text) > 10 Then
            Label7.Text = Strings.Left(TextBox1.Text, 13) 'SIMPAN KE LABEL7 SEBAGAI ACUAN PENCARIAN QUERY
            Label10.Text = Label7.Text + "000"
            Call connection()
            cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label10.Text & "%')", conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
                Dim UrutanKode As String
                Dim Hitung As Long
                rd = cmd.ExecuteReader
                rd.Read()
                If rd.HasRows Then
                    Label8.Text = rd("txn_id") 'DAPATIN ANGKA TERBESAR
                    If Len(Label8.Text) > 10 Then
                        Label8.Text = Strings.Left(Label8.Text, 13) 'SIMPAN KE LABEL9 + 1 SEBAGAI ACUAN FIFO
                        Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
                        UrutanKode = Label8.Text + Microsoft.VisualBasic.Right("000" & Hitung, 3)
                        Label9.Text = UrutanKode
                        rd.Close()
                        conn.Close()
                    End If
                    If TextBox1.Text <> Label9.Text Then
                        MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                        TextBox1.Clear()
                        Label7.Text = ""
                        Label8.Text = ""
                        Label9.Text = ""
                    ElseIf TextBox1.Text = Label9.Text Then
                        Call connection()
                        str = "SELECT * FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id='" & TextBox1.Text & "'"
                        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                        rd = cmd.ExecuteReader
                        rd.Read()
                        Try
                            If rd.HasRows Then
                                ComboBox1.Text = rd("topbott")
                                rd.Close()

                                cmd = New MySql.Data.MySqlClient.MySqlCommand
                                cmd.Connection = conn
                                str = "INSERT INTO `prodsys2_solder_paste_mgt_line_tbl`(`txn_id`, `topbott` , `line`, `start_time`, `status`, `pic`) VALUES (@txn_id,@topbott,@line,@start_time,@status,@pic)"
                                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox2.Text
                                cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                cmd.CommandText = str
                                Try
                                    cmd.ExecuteNonQuery()
                                    conn.Close()
                                    conn.Dispose()
                                Catch ex As Exception
                                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    rd.Close()
                                    conn.Close()
                                    conn.Dispose()
                                End Try
                            End If
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End Try

                        Dim sql As String
                        rd.Close()
                        Call connection()
                        sql = "UPDATE prodsys2_solder_paste_mgt_mixer_tbl SET status = @stat WHERE txn_id = @id"
                        cmd.Connection = conn
                        cmd.CommandText = sql
                        cmd.Parameters.AddWithValue("@id", TextBox1.Text)
                        cmd.Parameters.AddWithValue("@stat", "EMPTY")
                        Dim r As Integer
                        r = cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()

                        Call LoadDataCoolingCase()
                        Call LoadDataCheckRight()
                        Call LoadDataMixer()
                        Call LoadDataLine()

                        Call ShowCountCoolingCase()
                        Call ShowCountCheckRight()
                        Call ShowCountMixer()
                        Call ShowCountLine()

                        Call bersih()

                    ElseIf TextBox1.Text > Label9.Text Then
                        MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                        TextBox1.Clear()
                        Label7.Text = ""
                        Label8.Text = ""
                        Label9.Text = ""
                    End If
                Else
                    MsgBox("SOLDER PASTE BELUM DI MIXER!", vbInformation)
                    TextBox1.Clear()
                End If
            Else
                Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand
                cmd.Connection = conn
                str = "INSERT INTO `prodsys2_solder_paste_mgt_mixer_tbl`(`txn_id`, `status`) VALUES (@txn_id,@status)"
                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = Label10.Text
                cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "EMPTY"
                cmd.CommandText = str
                Try
                    cmd.ExecuteNonQuery()
                    conn.Close()
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try

                Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
                Dim UrutanKode As String
                Dim Hitung As Long
                rd = cmd.ExecuteReader
                rd.Read()
                If rd.HasRows Then
                    Label8.Text = rd("txn_id") 'DAPATIN ANGKA TERBESAR
                    If Len(Label8.Text) > 10 Then
                        Label8.Text = Strings.Left(Label8.Text, 13) 'SIMPAN KE LABEL9 + 1 SEBAGAI ACUAN FIFO
                        Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
                        UrutanKode = Label8.Text + Microsoft.VisualBasic.Right("000" & Hitung, 3)
                        Label9.Text = UrutanKode
                        rd.Close()
                        conn.Close()
                    End If
                    If TextBox1.Text <> Label9.Text Then
                        MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                        TextBox1.Clear()
                        Label7.Text = ""
                        Label8.Text = ""
                        Label9.Text = ""
                    ElseIf TextBox1.Text = Label9.Text Then
                        Call connection()
                        str = "SELECT * FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id='" & TextBox1.Text & "'"
                        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                        rd = cmd.ExecuteReader
                        rd.Read()
                        Try
                            If rd.HasRows Then
                                ComboBox1.Text = rd("topbott")
                                rd.Close()

                                cmd = New MySql.Data.MySqlClient.MySqlCommand
                                cmd.Connection = conn
                                str = "INSERT INTO `prodsys2_solder_paste_mgt_line_tbl`(`txn_id`, `topbott` , `line`, `start_time`, `status`, `pic`) VALUES (@txn_id,@topbott,@line,@start_time,@status,@pic)"
                                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox2.Text
                                cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                cmd.CommandText = str
                                Try
                                    cmd.ExecuteNonQuery()
                                    conn.Close()
                                    conn.Dispose()
                                Catch ex As Exception
                                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    rd.Close()
                                    conn.Close()
                                    conn.Dispose()
                                End Try
                            End If
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End Try

                        Dim sql As String
                        rd.Close()
                        Call connection()
                        sql = "UPDATE prodsys2_solder_paste_mgt_mixer_tbl SET status = @stat WHERE txn_id = @id"
                        cmd.Connection = conn
                        cmd.CommandText = sql
                        cmd.Parameters.AddWithValue("@id", TextBox1.Text)
                        cmd.Parameters.AddWithValue("@stat", "EMPTY")
                        Dim r As Integer
                        r = cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()

                        Call LoadDataCoolingCase()
                        Call LoadDataCheckRight()
                        Call LoadDataMixer()
                        Call LoadDataLine()

                        Call ShowCountCoolingCase()
                        Call ShowCountCheckRight()
                        Call ShowCountMixer()
                        Call ShowCountLine()

                        Call bersih()
                    ElseIf TextBox1.Text > Label9.Text Then
                        MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                        TextBox1.Clear()
                        Label7.Text = ""
                        Label8.Text = ""
                        Label9.Text = ""
                    End If
                Else
                    MsgBox("SOLDER PASTE BELUM DI MIXER!", vbInformation)
                    TextBox1.Clear()
                End If
            End If
        Else
            MsgBox("FORMAT BARCODE SALAH!", vbExclamation)
            TextBox1.Clear()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Len(TextBox1.Text) > 10 Then
                Label7.Text = Strings.Left(TextBox1.Text, 13) 'SIMPAN KE LABEL7 SEBAGAI ACUAN PENCARIAN QUERY
                Label10.Text = Label7.Text + "000"
                Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label10.Text & "%')", conn)
                rd = cmd.ExecuteReader
                rd.Read()
                If rd.HasRows Then
                    Call connection()
                    cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
                    Dim UrutanKode As String
                    Dim Hitung As Long
                    rd = cmd.ExecuteReader
                    rd.Read()
                    If rd.HasRows Then
                        Label8.Text = rd("txn_id") 'DAPATIN ANGKA TERBESAR
                        If Len(Label8.Text) > 10 Then
                            Label8.Text = Strings.Left(Label8.Text, 13) 'SIMPAN KE LABEL9 + 1 SEBAGAI ACUAN FIFO
                            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
                            UrutanKode = Label8.Text + Microsoft.VisualBasic.Right("000" & Hitung, 3)
                            Label9.Text = UrutanKode
                            rd.Close()
                            conn.Close()
                        End If
                        If TextBox1.Text <> Label9.Text Then
                            MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                            TextBox1.Clear()
                            Label7.Text = ""
                            Label8.Text = ""
                            Label9.Text = ""
                        ElseIf TextBox1.Text = Label9.Text Then
                            Call connection()
                            str = "SELECT * FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id='" & TextBox1.Text & "'"
                            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                            rd = cmd.ExecuteReader
                            rd.Read()
                            Try
                                If rd.HasRows Then
                                    ComboBox1.Text = rd("topbott")
                                    rd.Close()

                                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                                    cmd.Connection = conn
                                    str = "INSERT INTO `prodsys2_solder_paste_mgt_line_tbl`(`txn_id`, `topbott` , `line`, `start_time`, `status`, `pic`) VALUES (@txn_id,@topbott,@line,@start_time,@status,@pic)"
                                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                    cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox2.Text
                                    cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                    cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                    cmd.CommandText = str
                                    Try
                                        cmd.ExecuteNonQuery()
                                        conn.Close()
                                        conn.Dispose()
                                    Catch ex As Exception
                                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        rd.Close()
                                        conn.Close()
                                        conn.Dispose()
                                    End Try
                                End If
                            Catch ex As Exception
                                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            End Try

                            Dim sql As String
                            rd.Close()
                            Call connection()
                            sql = "UPDATE prodsys2_solder_paste_mgt_mixer_tbl SET status = @stat WHERE txn_id = @id"
                            cmd.Connection = conn
                            cmd.CommandText = sql
                            cmd.Parameters.AddWithValue("@id", TextBox1.Text)
                            cmd.Parameters.AddWithValue("@stat", "EMPTY")
                            Dim r As Integer
                            r = cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()

                            Call LoadDataCoolingCase()
                            Call LoadDataCheckRight()
                            Call LoadDataMixer()
                            Call LoadDataLine()

                            Call ShowCountCoolingCase()
                            Call ShowCountCheckRight()
                            Call ShowCountMixer()
                            Call ShowCountLine()

                            Call bersih()

                        ElseIf TextBox1.Text > Label9.Text Then
                            MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                            TextBox1.Clear()
                            Label7.Text = ""
                            Label8.Text = ""
                            Label9.Text = ""
                        End If
                    Else
                        MsgBox("SOLDER PASTE BELUM DI MIXER!", vbInformation)
                        TextBox1.Clear()
                    End If
                Else
                    Call connection()
                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_solder_paste_mgt_mixer_tbl`(`txn_id`, `status`) VALUES (@txn_id,@status)"
                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = Label10.Text
                    cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "EMPTY"
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try

                    Call connection()
                    cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
                    Dim UrutanKode As String
                    Dim Hitung As Long
                    rd = cmd.ExecuteReader
                    rd.Read()
                    If rd.HasRows Then
                        Label8.Text = rd("txn_id") 'DAPATIN ANGKA TERBESAR
                        If Len(Label8.Text) > 10 Then
                            Label8.Text = Strings.Left(Label8.Text, 13) 'SIMPAN KE LABEL9 + 1 SEBAGAI ACUAN FIFO
                            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
                            UrutanKode = Label8.Text + Microsoft.VisualBasic.Right("000" & Hitung, 3)
                            Label9.Text = UrutanKode
                            rd.Close()
                            conn.Close()
                        End If
                        If TextBox1.Text <> Label9.Text Then
                            MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                            TextBox1.Clear()
                            Label7.Text = ""
                            Label8.Text = ""
                            Label9.Text = ""
                        ElseIf TextBox1.Text = Label9.Text Then
                            Call connection()
                            str = "SELECT * FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id='" & TextBox1.Text & "'"
                            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                            rd = cmd.ExecuteReader
                            rd.Read()
                            Try
                                If rd.HasRows Then
                                    ComboBox1.Text = rd("topbott")
                                    rd.Close()

                                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                                    cmd.Connection = conn
                                    str = "INSERT INTO `prodsys2_solder_paste_mgt_line_tbl`(`txn_id`, `topbott` , `line`, `start_time`, `status`, `pic`) VALUES (@txn_id,@topbott,@line,@start_time,@status,@pic)"
                                    cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                    cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox2.Text
                                    cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                    cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                    cmd.CommandText = str
                                    Try
                                        cmd.ExecuteNonQuery()
                                        conn.Close()
                                        conn.Dispose()
                                    Catch ex As Exception
                                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        rd.Close()
                                        conn.Close()
                                        conn.Dispose()
                                    End Try
                                End If
                            Catch ex As Exception
                                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            End Try

                            Dim sql As String
                            rd.Close()
                            Call connection()
                            sql = "UPDATE prodsys2_solder_paste_mgt_mixer_tbl SET status = @stat WHERE txn_id = @id"
                            cmd.Connection = conn
                            cmd.CommandText = sql
                            cmd.Parameters.AddWithValue("@id", TextBox1.Text)
                            cmd.Parameters.AddWithValue("@stat", "EMPTY")
                            Dim r As Integer
                            r = cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()

                            Call LoadDataCoolingCase()
                            Call LoadDataCheckRight()
                            Call LoadDataMixer()
                            Call LoadDataLine()

                            Call ShowCountCoolingCase()
                            Call ShowCountCheckRight()
                            Call ShowCountMixer()
                            Call ShowCountLine()

                            Call bersih()
                        ElseIf TextBox1.Text > Label9.Text Then
                            MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                            TextBox1.Clear()
                            Label7.Text = ""
                            Label8.Text = ""
                            Label9.Text = ""
                        End If
                    Else
                        MsgBox("SOLDER PASTE BELUM DI MIXER!", vbInformation)
                        TextBox1.Clear()
                    End If
                End If
            Else
                    MsgBox("FORMAT BARCODE SALAH!", vbExclamation)
                TextBox1.Clear()
            End If
        End If
    End Sub
End Class