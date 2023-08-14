Imports MySql.Data.MySqlClient

Public Class FormCheckRight

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

    Sub CheckFIFO()
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE txn_id LIKE '" & Label7.Text & "%')", conn)
            Dim UrutanKode As String
            Dim Hitung As Long
            rd = cmd.ExecuteReader
            rd.Read()
        If rd.HasRows Then
            Label7.Text = rd("txn_id")
            If Len(Label7.Text) > 10 Then
                Label7.Text = Strings.Left(Label7.Text, 13)
            End If
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            UrutanKode = Label7.Text + Microsoft.VisualBasic.Right("000" & Hitung, 3)
            Label8.Text = UrutanKode
            rd.Close()
            conn.Close()
        Else
            MsgBox("data tidak ada")
        End If
    End Sub

    Private Sub FormCheckRight_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call bersih()

        For i As Integer = 1 To 20
            ComboBox2.Items.Add(i.ToString())
        Next
        TextBox4.Text = FormMain.LabelNama.Text
        TextBox1.Select()

        Label7.Text = ""
        Label8.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Len(TextBox1.Text) > 10 Then
            Label7.Text = Strings.Left(TextBox1.Text, 13) 'SIMPAN KE LABEL7 SEBAGAI ACUAN PENCARIAN QUERY
            Call connection()
            cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
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
                    'MsgBox("FIFO!", MsgBoxStyle.Exclamation)
                    'TextBox1.Clear()
                    'Label7.Text = ""
                    'Label8.Text = ""
                    'Label9.Text = ""
                    Call connection()
                    str = "SELECT * FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id='" & TextBox1.Text & "' AND status = 'AVAILABLE'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    rd.Read()
                    Try
                        If rd.HasRows Then
                            TextBox2.Text = rd("maker")
                            ComboBox1.Text = rd("topbott")
                            rd.Close()

                            If ComboBox2.Text = "" Then
                                MsgBox("HOLE BELUM DI PILIH!", vbInformation)
                            Else
                                'conn.Open()
                                str = "SELECT hole FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE hole='" & ComboBox2.Text & "' AND status = 'AVAILABLE'"
                                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                                rd = cmd.ExecuteReader
                                rd.Read()
                                Try
                                    If rd.HasRows Then
                                        MsgBox("HOLE [" & ComboBox2.Text & "] SUDAH TERPAKAI!", vbExclamation)
                                    Else
                                        For Baris As Integer = 0 To FormSolderPasteMGT.DataGridView1.Rows.Count - 1
                                            Dim sql As String
                                            rd.Close()
                                            sql = "UPDATE prodsys2_solder_paste_mgt_coolingcase_tbl SET run_time = @run_time, end_time = @end_time, status = @status WHERE txn_id = @txn_id"
                                            cmd.Connection = conn
                                            cmd.CommandText = sql
                                            cmd.Parameters.AddWithValue("@txn_id", TextBox1.Text)
                                            cmd.Parameters.AddWithValue("@run_time", FormSolderPasteMGT.DataGridView1.Rows(Baris).Cells(5).Value)
                                            cmd.Parameters.AddWithValue("@end_time", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                                            cmd.Parameters.AddWithValue("@status", "EMPTY")
                                            Dim r As Integer
                                            r = cmd.ExecuteNonQuery()
                                            cmd.Parameters.Clear()
                                        Next
                                        Call connection()
                                        rd.Close()
                                        cmd = New MySql.Data.MySqlClient.MySqlCommand
                                        cmd.Connection = conn
                                        str = "INSERT INTO `prodsys2_solder_paste_mgt_checkright_tbl`(`txn_id`, `maker`, `n_type`, `topbott` , `hole`, `start_time`, `status`, `pic`) VALUES (@txn_id,@maker,@n_type,@topbott,@hole,@start_time,@status,@pic)"
                                        cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                        cmd.Parameters.Add("@maker", MySqlDbType.VarChar).Value = TextBox2.Text
                                        cmd.Parameters.Add("@n_type", MySqlDbType.VarChar).Value = TextBox3.Text
                                        cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                        cmd.Parameters.Add("@hole", MySqlDbType.VarChar).Value = ComboBox2.Text
                                        cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                        cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                        cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = TextBox4.Text
                                        cmd.CommandText = str
                                        Try
                                            cmd.ExecuteNonQuery()
                                            conn.Close()
                                            conn.Dispose()

                                            Call LoadDataCoolingCase()
                                            Call LoadDataCheckRight()
                                            Call LoadDataMixer()
                                            Call LoadDataLine()

                                            Call ShowCountCoolingCase()
                                            Call ShowCountCheckRight()
                                            Call ShowCountMixer()
                                            Call ShowCountLine()

                                            Call bersih()
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
                            End If
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        Else
                            MsgBox("SOLDER PASTE TIDAK ADA DI COOLING CASE ATAU BELUM TERDAFTAR!", vbInformation)
                            Call bersih()
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End If
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                ElseIf TextBox1.Text > Label9.Text Then
                    MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                    TextBox1.Clear()
                    Label7.Text = ""
                    Label8.Text = ""
                    Label9.Text = ""
                End If
            Else
                MsgBox("DATA TIDAK ADA!", vbInformation)
                TextBox1.Clear()
            End If
        Else
            MsgBox("FORMAT BARCODE SALAH!", vbExclamation)
            TextBox1.Clear()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Len(TextBox1.Text) > 10 Then
                Label7.Text = Strings.Left(TextBox1.Text, 13) 'SIMPAN KE LABEL7 SEBAGAI ACUAN PENCARIAN QUERY
                Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT txn_id FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id LIKE '" & Label7.Text & "%' AND status ='EMPTY')", conn)
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
                        'MsgBox("FIFO!", MsgBoxStyle.Exclamation)
                        'TextBox1.Clear()
                        'Label7.Text = ""
                        'Label8.Text = ""
                        'Label9.Text = ""
                        Call connection()
                        str = "SELECT * FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE txn_id='" & TextBox1.Text & "' AND status = 'AVAILABLE'"
                        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                        rd = cmd.ExecuteReader
                        rd.Read()
                        Try
                            If rd.HasRows Then
                                TextBox2.Text = rd("maker")
                                ComboBox1.Text = rd("topbott")
                                rd.Close()

                                If ComboBox2.Text = "" Then
                                    MsgBox("HOLE BELUM DI PILIH!", vbInformation)
                                Else
                                    'conn.Open()
                                    str = "SELECT hole FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE hole='" & ComboBox2.Text & "' AND status = 'AVAILABLE'"
                                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                                    rd = cmd.ExecuteReader
                                    rd.Read()
                                    Try
                                        If rd.HasRows Then
                                            MsgBox("HOLE [" & ComboBox2.Text & "] SUDAH TERPAKAI!", vbExclamation)
                                        Else
                                            For Baris As Integer = FormSolderPasteMGT.DataGridView1.Rows.Count - 1 To 0 Step -1
                                                rd.Close()
                                                Dim cmd As New MySqlCommand("UPDATE prodsys2_solder_paste_mgt_coolingcase_tbl SET run_time = @run_time ,end_time = @end_time ,status = @status WHERE txn_id = @txn_id AND status='AVAILABLE'", conn)
                                                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                                cmd.Parameters.Add("@run_time", MySqlDbType.VarChar).Value = FormSolderPasteMGT.DataGridView1.Rows(Baris).Cells(5).Value
                                                cmd.Parameters.Add("@end_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
                                                cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "EMPTY"
                                                cmd.ExecuteNonQuery()
                                            Next
                                            'Call connection()
                                            cmd = New MySql.Data.MySqlClient.MySqlCommand
                                            cmd.Connection = conn
                                            str = "INSERT INTO `prodsys2_solder_paste_mgt_checkright_tbl` (`txn_id`, `maker`, `topbott` , `hole`, `start_time`, `status`, `pic`) VALUES (@txn_id,@maker,@topbott,@hole,@start_time,@status,@pic)"
                                            cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                                            cmd.Parameters.Add("@maker", MySqlDbType.VarChar).Value = TextBox2.Text
                                            cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = ComboBox1.Text
                                            cmd.Parameters.Add("@hole", MySqlDbType.VarChar).Value = ComboBox2.Text
                                            cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                            cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                                            cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = TextBox4.Text
                                            cmd.CommandText = str
                                            Try
                                                cmd.ExecuteNonQuery()
                                                Call LoadDataCoolingCase()
                                                Call LoadDataCheckRight()
                                                Call LoadDataMixer()
                                                Call LoadDataLine()

                                                Call ShowCountCoolingCase()
                                                Call ShowCountCheckRight()
                                                Call ShowCountMixer()
                                                Call ShowCountLine()

                                                Call bersih()
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
                                End If
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            Else
                                MsgBox("SOLDER PASTE TIDAK ADA DI COOLING CASE ATAU BELUM TERDAFTAR!", vbInformation)
                                Call bersih()
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            End If
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End Try
                    ElseIf TextBox1.Text > Label9.Text Then
                        MsgBox("TIDAK FIFO!", MsgBoxStyle.Exclamation)
                        TextBox1.Clear()
                        Label7.Text = ""
                        Label8.Text = ""
                        Label9.Text = ""
                    End If
                Else
                    MsgBox("DATA TIDAK ADA!", vbInformation)
                    TextBox1.Clear()
                End If
            Else
                MsgBox("FORMAT BARCODE SALAH!", vbExclamation)
                TextBox1.Clear()
            End If
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox1.Clear()
        TextBox1.Select()
    End Sub
End Class