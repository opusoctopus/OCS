Imports MySql.Data.MySqlClient

Public Class FormAttachMetalmask

    Sub Bersih()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox17.Clear()
        TextBox18.Clear()
        TextBox19.Clear()

        TextBox13.Select()

        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""

        ComboBox1.SelectedItem = Nothing
        ComboBox5.SelectedItem = Nothing
    End Sub

    Sub LoadDataInspection()
        Try
            str = "SELECT kode_stencil as KodeStencil, 
                    pn_pcb as PartNumber, 
                    prod_qty as Qty, 
                    cleaning_status as Cleaned, 
                    tension_pa as TensionP1, 
                    tension_pb as TensionP2, 
                    tension_pc as TensionP3, 
                    tension_pd as TensionP4, 
                    tension_pe as TensionP5, 
                    pic as PIC, 
                    created as Created
                   FROM prodsys2_stencil_trans_in_tbl ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormStencilMGT.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInspection()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT line as Line, 
                    kode_stencil as KodeStencil, 
                    pn_pcb as PartNumber,       
                    pic as PIC, 
                    created as Created 
                   FROM prodsys2_stencil_trans_out_tbl WHERE status='AVAILABLE' ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormStencilMGT.DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInLine()
    End Sub

    Sub BuatKolomInspection()
        FormStencilMGT.DataGridView2.Columns(0).Width = 120
        FormStencilMGT.DataGridView2.Columns(1).Width = 110
        FormStencilMGT.DataGridView2.Columns(2).Width = 60
        FormStencilMGT.DataGridView2.Columns(3).Width = 60
        FormStencilMGT.DataGridView2.Columns(4).Width = 80
        FormStencilMGT.DataGridView2.Columns(5).Width = 80
        FormStencilMGT.DataGridView2.Columns(6).Width = 80
        FormStencilMGT.DataGridView2.Columns(7).Width = 80
        FormStencilMGT.DataGridView2.Columns(8).Width = 80
        FormStencilMGT.DataGridView2.Columns(9).Width = 80
        FormStencilMGT.DataGridView2.Columns(10).Width = 130
    End Sub

    Sub BuatKolomInLine()
        FormStencilMGT.DataGridView3.Columns(0).Width = 80
        FormStencilMGT.DataGridView3.Columns(1).Width = 120
        FormStencilMGT.DataGridView3.Columns(2).Width = 120
        FormStencilMGT.DataGridView3.Columns(3).Width = 80
        FormStencilMGT.DataGridView3.Columns(4).Width = 130
    End Sub

    Private Sub FormAttachMetalmask_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Call Bersih()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox17.Clear()
        TextBox18.Clear()
        TextBox19.Clear()

        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""

        TextBox13.Select()

        ComboBox5.SelectedItem = Nothing
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call Bersih()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        If TextBox13.Text = "" Or TextBox14.Text = "" Or TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or ComboBox1.SelectedItem = Nothing Or ComboBox5.SelectedItem = Nothing Then
            MsgBox("Data belum lengkap!", MsgBoxStyle.Critical)
        ElseIf Label2.Text = "NG" Or Label3.Text = "NG" Or Label4.Text = "NG" Or Label5.Text = "NG" Or Label6.Text = "NG" Then
            MsgBox("RESULT TENSION CHECK STATUS [NG]", vbCritical)
        Else
            If Val(Label7.Text) < 40000 Then
                If ComboBox1.Text = "BOTTOM" Then
                    rd.Close()
                    Dim cmd As New MySqlCommand("UPDATE prodsys2_stencil_inventory_tbl SET usages = @usages WHERE kode_stencil = @kode_stencil", conn)
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@usages", MySqlDbType.Int64).Value = Label7.Text
                    cmd.ExecuteNonQuery()

                    Dim sql2 As String
                    sql2 = "UPDATE prodsys2_stencil_trans_out_tbl SET status = @stat WHERE kode_stencil='" & TextBox13.Text & "' AND status='AVAILABLE'"
                    cmd.Connection = conn
                    cmd.CommandText = sql2
                    cmd.Parameters.AddWithValue("@stat", "EMPTY")
                    Dim r2 As Integer
                    r2 = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Dim cmd2 As New MySqlCommand("UPDATE prodsys2_stencil_rack SET status = @stat WHERE kode_stencil = @kode_stencil", conn)
                    cmd2.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd2.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "RACK"
                    cmd2.ExecuteNonQuery()

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_trans_in_tbl` (`kode_stencil`, `pn_pcb`, `layer`, `line`, `prod_qty`, `cleaning_status`, `tension_pa`, `tension_pb`, `tension_pc`, `tension_pd`, `tension_pe`, `result`, `remark`, `pic`, `created`)
VALUES (@kode_stencil,@pn_pcb,@layer,@line, @prod_qty, @cleaning_status, @tension_pa, @tension_pb, @tension_pc, @tension_pd, @tension_pe, @result, @remark, @pic, @created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = TextBox14.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = ComboBox1.Text
                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox5.Text
                    cmd.Parameters.Add("@prod_qty", MySqlDbType.VarChar).Value = TextBox17.Text
                    cmd.Parameters.Add("@cleaning_status", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@tension_pa", MySqlDbType.VarChar).Value = TextBox8.Text
                    cmd.Parameters.Add("@tension_pb", MySqlDbType.VarChar).Value = TextBox9.Text
                    cmd.Parameters.Add("@tension_pc", MySqlDbType.VarChar).Value = TextBox10.Text
                    cmd.Parameters.Add("@tension_pd", MySqlDbType.VarChar).Value = TextBox11.Text
                    cmd.Parameters.Add("@tension_pe", MySqlDbType.VarChar).Value = TextBox12.Text
                    cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = DBNull.Value
                    cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                    Call Bersih()
                    Call LoadDataInLine()
                    Call LoadDataInspection()
                ElseIf ComboBox1.Text = "TOP" Then
                    rd.Close()
                    Dim cmd As New MySqlCommand("UPDATE prodsys2_stencil_inventory_tbl SET usages = @usages WHERE kode_stencil = @kode_stencil", conn)
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@usages", MySqlDbType.Int64).Value = Val(Label7.Text) + Val(TextBox17.Text)
                    cmd.ExecuteNonQuery()

                    Dim sql2 As String
                    sql2 = "UPDATE prodsys2_stencil_trans_out_tbl SET status = @stat WHERE kode_stencil='" & TextBox13.Text & "' AND status='AVAILABLE'"
                    cmd.Connection = conn
                    cmd.CommandText = sql2
                    cmd.Parameters.AddWithValue("@stat", "EMPTY")
                    Dim r2 As Integer
                    r2 = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Dim cmd2 As New MySqlCommand("UPDATE prodsys2_stencil_rack SET status = @stat WHERE kode_stencil = @kode_stencil", conn)
                    cmd2.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd2.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "RACK"
                    cmd2.ExecuteNonQuery()

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_trans_in_tbl` (`kode_stencil`, `pn_pcb`, `layer`, `line`, `prod_qty`, `cleaning_status`, `tension_pa`, `tension_pb`, `tension_pc`, `tension_pd`, `tension_pe`, `result`, `remark`, `pic`, `created`)
VALUES (@kode_stencil,@pn_pcb,@layer,@line, @prod_qty, @cleaning_status, @tension_pa, @tension_pb, @tension_pc, @tension_pd, @tension_pe, @result, @remark, @pic, @created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = TextBox14.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = ComboBox1.Text
                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox5.Text
                    cmd.Parameters.Add("@prod_qty", MySqlDbType.VarChar).Value = TextBox17.Text
                    cmd.Parameters.Add("@cleaning_status", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@tension_pa", MySqlDbType.VarChar).Value = TextBox8.Text
                    cmd.Parameters.Add("@tension_pb", MySqlDbType.VarChar).Value = TextBox9.Text
                    cmd.Parameters.Add("@tension_pc", MySqlDbType.VarChar).Value = TextBox10.Text
                    cmd.Parameters.Add("@tension_pd", MySqlDbType.VarChar).Value = TextBox11.Text
                    cmd.Parameters.Add("@tension_pe", MySqlDbType.VarChar).Value = TextBox12.Text
                    cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = DBNull.Value
                    cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                    Call Bersih()
                    Call LoadDataInLine()
                    Call LoadDataInspection()
                End If
            ElseIf Val(Label7.Text) < 50000 Then
                If ComboBox1.Text = "BOTTOM" Then
                    MsgBox("[SYSTEM-WARNING] - PEMAKAIAN METALMASK INI HAMPIR HABIS, SEGERA LAKUKAN TINDAKAN PREVENTIF!", MsgBoxStyle.Exclamation)

                    Dim cmd As New MySqlCommand("UPDATE prodsys2_stencil_inventory_tbl SET usages = @usages WHERE kode_stencil = @kode_stencil", conn)
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@usages", MySqlDbType.Int64).Value = Label7.Text
                    cmd.ExecuteNonQuery()

                    Dim sql2 As String
                    sql2 = "UPDATE prodsys2_stencil_trans_out_tbl SET status = @stat WHERE kode_stencil='" & TextBox13.Text & "' AND status='AVAILABLE'"
                    cmd.Connection = conn
                    cmd.CommandText = sql2
                    cmd.Parameters.AddWithValue("@stat", "EMPTY")
                    Dim r2 As Integer
                    r2 = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Dim cmd2 As New MySqlCommand("UPDATE prodsys2_stencil_rack SET status = @stat WHERE kode_stencil = @kode_stencil", conn)
                    cmd2.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd2.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "RACK"
                    cmd2.ExecuteNonQuery()

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_trans_in_tbl` (`kode_stencil`, `pn_pcb`, `layer` , `line`, `prod_qty`, `cleaning_status`, `tension_pa`, `tension_pb`, `tension_pc`, `tension_pd`, `tension_pe`, `remark`, `pic`, `created`)
VALUES (@kode_stencil,@pn_pcb,@layer,@line, @prod_qty, @cleaning_status, @tension_pa, @tension_pb, @tension_pc, @tension_pd, @tension_pe, @result, @pic, @created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = TextBox14.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = ComboBox1.Text
                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox5.Text
                    cmd.Parameters.Add("@prod_qty", MySqlDbType.VarChar).Value = TextBox17.Text
                    cmd.Parameters.Add("@cleaning_status", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@tension_pa", MySqlDbType.VarChar).Value = TextBox8.Text
                    cmd.Parameters.Add("@tension_pb", MySqlDbType.VarChar).Value = TextBox9.Text
                    cmd.Parameters.Add("@tension_pc", MySqlDbType.VarChar).Value = TextBox10.Text
                    cmd.Parameters.Add("@tension_pd", MySqlDbType.VarChar).Value = TextBox11.Text
                    cmd.Parameters.Add("@tension_pe", MySqlDbType.VarChar).Value = TextBox12.Text
                    cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                    Call Bersih()
                    Call LoadDataInLine()
                    Call LoadDataInspection()
                ElseIf ComboBox1.Text = "TOP" Then
                    MsgBox("[SYSTEM-WARNING] - PEMAKAIAN METALMASK INI HAMPIR HABIS, SEGERA LAKUKAN TINDAKAN PREVENTIF!", MsgBoxStyle.Exclamation)

                    Dim cmd As New MySqlCommand("UPDATE prodsys2_stencil_inventory_tbl SET usages = @usages WHERE kode_stencil = @kode_stencil", conn)
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@usages", MySqlDbType.Int64).Value = Val(Label7.Text) + Val(TextBox17.Text)
                    cmd.ExecuteNonQuery()

                    Dim sql2 As String
                    sql2 = "UPDATE prodsys2_stencil_trans_out_tbl SET status = @stat WHERE kode_stencil='" & TextBox13.Text & "' AND status='AVAILABLE'"
                    cmd.Connection = conn
                    cmd.CommandText = sql2
                    cmd.Parameters.AddWithValue("@stat", "EMPTY")
                    Dim r2 As Integer
                    r2 = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Dim cmd2 As New MySqlCommand("UPDATE prodsys2_stencil_rack SET status = @stat WHERE kode_stencil = @kode_stencil", conn)
                    cmd2.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd2.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "RACK"
                    cmd2.ExecuteNonQuery()

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_trans_in_tbl` (`kode_stencil`, `pn_pcb`, `layer` , `line`, `prod_qty`, `cleaning_status`, `tension_pa`, `tension_pb`, `tension_pc`, `tension_pd`, `tension_pe`, `remark`, `pic`, `created`)
VALUES (@kode_stencil,@pn_pcb,@layer,@line, @prod_qty, @cleaning_status, @tension_pa, @tension_pb, @tension_pc, @tension_pd, @tension_pe, @result, @pic, @created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = TextBox14.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = ComboBox1.Text
                    cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox5.Text
                    cmd.Parameters.Add("@prod_qty", MySqlDbType.VarChar).Value = TextBox17.Text
                    cmd.Parameters.Add("@cleaning_status", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@tension_pa", MySqlDbType.VarChar).Value = TextBox8.Text
                    cmd.Parameters.Add("@tension_pb", MySqlDbType.VarChar).Value = TextBox9.Text
                    cmd.Parameters.Add("@tension_pc", MySqlDbType.VarChar).Value = TextBox10.Text
                    cmd.Parameters.Add("@tension_pd", MySqlDbType.VarChar).Value = TextBox11.Text
                    cmd.Parameters.Add("@tension_pe", MySqlDbType.VarChar).Value = TextBox12.Text
                    cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox18.Text
                    cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                    Call Bersih()
                    Call LoadDataInLine()
                    Call LoadDataInspection()
                End If
            ElseIf Val(Label7.Text) >= 50000 Then
                    MsgBox("[SYSTEM-ALERT] - TANSAKSI DIBATALKAN, PEMAKAIAN METALMASK INI SUDAH MELEBIHI BATAS MAXIMUM PEMAKAIAN!", MsgBoxStyle.Critical)
                Call Bersih()
            End If
        End If
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox10.Text) < 0.5 Or Val(TextBox10.Text) > 0.8 Then
                TextBox10.ForeColor = Color.FromArgb(200, 5, 83)
                Label4.ForeColor = Color.FromArgb(200, 5, 83)
                Label4.Text = "NG"
                MsgBox("SELF CHECK: TENSION 3 [NG]", vbCritical)
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox11.Select()
            ElseIf Val(TextBox10.Text) >= 0.5 Or Val(TextBox10.Text) <= 0.8 Then
                TextBox10.ForeColor = Color.Black
                Label4.ForeColor = Color.Green
                Label4.Text = "OK"
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox11.Select()
            End If
        End If
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox11.Text) < 0.5 Or Val(TextBox11.Text) > 0.8 Then
                TextBox11.ForeColor = Color.FromArgb(200, 5, 83)
                Label5.ForeColor = Color.FromArgb(200, 5, 83)
                Label5.Text = "NG"
                MsgBox("SELF CHECK: TENSION 4 [NG]", vbCritical)
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox12.Select()
            ElseIf Val(TextBox11.Text) >= 0.5 Or Val(TextBox11.Text) <= 0.8 Then
                TextBox11.ForeColor = Color.Black
                Label5.ForeColor = Color.Green
                Label5.Text = "OK"
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox12.Select()
            End If
        End If
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox12.Text) < 0.5 Or Val(TextBox12.Text) > 0.8 Then
                TextBox12.ForeColor = Color.FromArgb(200, 5, 83)
                Label6.ForeColor = Color.FromArgb(200, 5, 83)
                Label6.Text = "NG"
                MsgBox("SELF CHECK: TENSION 5 [NG]", vbCritical)
                e.Handled = True
                SendKeys.SendWait("{TAB}")
            ElseIf Val(TextBox12.Text) >= 0.5 Or Val(TextBox12.Text) <= 0.8 Then
                TextBox12.ForeColor = Color.Black
                Label2.ForeColor = Color.Green
                Label6.Text = "OK"
                e.Handled = True
                SendKeys.SendWait("{TAB}")
            End If
        End If
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox8.Text) < 0.5 Or Val(TextBox8.Text) > 0.8 Then
                TextBox8.ForeColor = Color.FromArgb(200, 5, 83)
                Label2.ForeColor = Color.FromArgb(200, 5, 83)
                Label2.Text = "NG"
                MsgBox("SELF CHECK: TENSION 1 [NG]", vbCritical)
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox9.Select()
            ElseIf Val(TextBox8.Text) >= 0.5 Or Val(TextBox8.Text) <= 0.8 Then
                TextBox8.ForeColor = Color.Black
                Label2.ForeColor = Color.Green
                Label2.Text = "OK"
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox9.Select()
            End If
        End If
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox9.Text) < 0.5 Or Val(TextBox9.Text) > 0.8 Then
                TextBox9.ForeColor = Color.FromArgb(200, 5, 83)
                Label3.ForeColor = Color.FromArgb(200, 5, 83)
                Label3.Text = "NG"
                MsgBox("SELF CHECK: TENSION 2 [NG]", vbCritical)
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox10.Select()
            ElseIf Val(TextBox9.Text) >= 0.5 Or Val(TextBox9.Text) <= 0.8 Then
                TextBox9.ForeColor = Color.Black
                Label3.ForeColor = Color.Green
                Label3.Text = "OK"
                e.Handled = True
                SendKeys.SendWait("{TAB}")
                TextBox10.Select()
            End If
        End If
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If e.KeyChar = Chr(13) Then
            Call connection()
            str = "SELECT * FROM prodsys2_stencil_trans_out_tbl WHERE kode_stencil='" & TextBox13.Text & "' AND status='AVAILABLE'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            Try
                If rd.HasRows Then
                    While rd.Read
                        TextBox13.Text = rd("kode_stencil")
                        TextBox14.Text = rd("pn_pcb")
                        ComboBox1.Text = rd("layer")
                        ComboBox5.Text = rd("line")
                    End While

                    rd.Close()
                    str = "SELECT date as DATE, pcbname as PCB, COUNT(pcbname) as QTY FROM prodsys2_spi_log_tbl WHERE line ='" & ComboBox5.Text & "' AND pcbname='" & TextBox14.Text & "' AND date = DATE(NOW()) GROUP BY date, pcbname ORDER BY date ASC"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    Try
                        While rd.Read
                            TextBox17.Text = rd("QTY")
                        End While
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try

                    rd.Close()
                    str = "SELECT * FROM prodsys2_stencil_inventory_tbl WHERE kode_stencil='" & TextBox13.Text & "'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    Try
                        While rd.Read
                            Label7.Text = rd("usages")
                            If Val(Label7.Text) < 40000 Then
                                TextBox18.Select()
                            ElseIf Val(Label7.Text) < 50000 Then
                                MsgBox("[SYSTEM-WARNING] - PEMAKAIAN METALMASK INI HAMPIR HABIS, SEGERA LAKUKAN TINDAKAN PREVENTIVE!", MsgBoxStyle.Exclamation)
                                TextBox18.Select()
                            ElseIf Val(Label7.Text) >= 50000 Then
                                MsgBox("[SYSTEM-ALERT] - TANSAKSI DIBATALKAN, PEMAKAIAN METALMASK INI SUDAH MELEBIHI BATAS MAXIMUM PEMAKAIAN!", MsgBoxStyle.Critical)
                                Call Bersih()
                            End If
                        End While
                        rd.Close()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                Else
                    MsgBox("METALMASK INI TIDAK ADA DI LINE!", vbInformation)
                    Call Bersih()
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
        End If
    End Sub
End Class