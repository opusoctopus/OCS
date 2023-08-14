Imports MySql.Data.MySqlClient

Public Class FormStencilSelfCheckIncoming

    Dim arrImage1() As Byte
    Dim arrImage2() As Byte

    Sub Bersih()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()

        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""

        PictureBox2.Image = Nothing
        PictureBox3.Image = Nothing
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
                    DATE_FORMAT(STR_TO_DATE(created,'%d/%m/%Y'), '%Y/%m/%d') as Created
                   FROM prodsys2_stencil_trans_in_tbl ORDER BY Created DESC"
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FormStencilInventory.TextBox1.Text = "" Or
                FormStencilInventory.TextBox2.Text = "" Or
                FormStencilInventory.TextBox3.Text = "" Or
                FormStencilInventory.TextBox4.Text = "" Or
                FormStencilInventory.TextBox5.Text = "" Or
                FormStencilInventory.TextBox6.Text = "" Or
                FormStencilInventory.TextBox7.Text = "" Or
                FormStencilInventory.TextBox9.Text = "" Or
                FormStencilInventory.TextBox10.Text = "" Or
                FormStencilInventory.TextBox11.Text = "" Or
                FormStencilInventory.TextBox12.Text = "" Or
                FormStencilInventory.ComboBox1.Text = "" Or
                FormStencilInventory.ComboBox2.Text = "" Or
                TextBox8.Text = "" Or
                TextBox9.Text = "" Or
                TextBox10.Text = "" Or
                TextBox11.Text = "" Or
                TextBox12.Text = "" Or
                PictureBox2.Image Is Nothing Then
            MsgBox("Data belum lengkap!", MsgBoxStyle.Information)
        ElseIf Label2.Text = "NG" Or Label3.Text = "NG" Or Label4.Text = "NG" Or Label5.Text = "NG" Or Label6.Text = "NG" Then
            MsgBox("RESULT TENSION FROM SELF CHECK STATUS [NG]", vbCritical)
        Else
            'SIMPAN DATA
            Call connection()
            str = "SELECT * FROM prodsys2_stencil_inventory_tbl WHERE kode_stencil='" & FormStencilInventory.TextBox10.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    MsgBox("KODE STENCIL DUPLIKAT!", vbExclamation)
                Else
                    rd.Close()

                    Dim mstream1 As New System.IO.MemoryStream()
                    'Dim mstream2 As New System.IO.MemoryStream()
                    PictureBox2.Image.Save(mstream1, System.Drawing.Imaging.ImageFormat.Jpeg)
                    'PictureBox3.Image.Save(mstream2, System.Drawing.Imaging.ImageFormat.Jpeg)
                    arrImage1 = mstream1.GetBuffer()
                    'arrImage2 = mstream2.GetBuffer()
                    Dim FileSize1 As UInt32
                    'Dim FileSize2 As UInt32
                    FileSize1 = mstream1.Length
                    'FileSize2 = mstream2.Length
                    mstream1.Close()
                    'mstream2.Close()

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_inventory_tbl`(`kode_stencil`, `maker_supply`, `pn_pcb`, `layer`, `pabrication_date`, `incoming_date`, `location`, `lifetime`, `usages`, `created`) 
                            VALUES (@kode_stencil,@maker_supply,@pn_pcb,@layer,@pabrication_date,@incoming_date,@location,@lifetime,@usages,@created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox10.Text
                    cmd.Parameters.Add("@maker_supply", MySqlDbType.VarChar).Value = FormStencilInventory.ComboBox1.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox9.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = FormStencilInventory.ComboBox2.Text
                    cmd.Parameters.Add("@pabrication_date", MySqlDbType.VarChar).Value = FormStencilInventory.DateTimePicker1.Value.Date
                    cmd.Parameters.Add("@incoming_date", MySqlDbType.VarChar).Value = FormStencilInventory.DateTimePicker2.Value.Date
                    cmd.Parameters.Add("@location", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox11.Text
                    cmd.Parameters.Add("@lifetime", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox12.Text
                    cmd.Parameters.Add("@usages", MySqlDbType.VarChar).Value = "0"
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_audit_data_tbl`(`kode_stencil`, `pn_pcb`, `thickness`, `version`, `layer`, `maker_tension_point_a`, `maker_tension_point_b`, `maker_tension_point_c`, `maker_tension_point_d`, `maker_tension_point_e`, `self_tension_point_a`, `self_tension_point_b`, `self_tension_point_c`, `self_tension_point_d`, `self_tension_point_e`, `evidence1`, `evidence2`, `created`)
                            VALUES (@kode_stencil,@pn_pcb,@thickness,@version,@layer,@maker_tension_point_a,@maker_tension_point_b,@maker_tension_point_c,@maker_tension_point_d,@maker_tension_point_e,@self_tension_point_a,@self_tension_point_b,@self_tension_point_c,@self_tension_point_d,@self_tension_point_e,@evidence1,@evidence2,@created)"
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox10.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox9.Text
                    cmd.Parameters.Add("@thickness", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox1.Text
                    cmd.Parameters.Add("@version", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox2.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = FormStencilInventory.ComboBox2.Text
                    cmd.Parameters.Add("@maker_tension_point_a", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox3.Text
                    cmd.Parameters.Add("@maker_tension_point_b", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox4.Text
                    cmd.Parameters.Add("@maker_tension_point_c", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox5.Text
                    cmd.Parameters.Add("@maker_tension_point_d", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox6.Text
                    cmd.Parameters.Add("@maker_tension_point_e", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox7.Text
                    cmd.Parameters.Add("@self_tension_point_a", MySqlDbType.VarChar).Value = TextBox8.Text
                    cmd.Parameters.Add("@self_tension_point_b", MySqlDbType.VarChar).Value = TextBox9.Text
                    cmd.Parameters.Add("@self_tension_point_c", MySqlDbType.VarChar).Value = TextBox10.Text
                    cmd.Parameters.Add("@self_tension_point_d", MySqlDbType.VarChar).Value = TextBox11.Text
                    cmd.Parameters.Add("@self_tension_point_e", MySqlDbType.VarChar).Value = TextBox12.Text
                    cmd.Parameters.Add("@evidence1", MySqlDbType.LongBlob).Value = arrImage1
                    'cmd.Parameters.Add("@evidence2", MySqlDbType.VarChar).Value = arrImage2
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try

                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_stencil_rack` (`location`, `kode_stencil`, `pn_pcb`, `layer`, `status`, `created`) VALUES (@location,@kode_stencil,@pn_pcb,@layer,@status,@created)"
                    cmd.Parameters.Add("@location", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox11.Text
                    cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox10.Text
                    cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = FormStencilInventory.TextBox9.Text
                    cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = FormStencilInventory.ComboBox2.Text
                    cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "RACK"
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
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
            rd.Close()
            conn.Close()
            conn.Dispose()

            Call LoadDataInLine()
            Call LoadDataInspection()
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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        FormStencilEvidence1.ShowDialog()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        FormStencilEvidence2.ShowDialog()
    End Sub

    Private Sub FormStencilSelfCheckIncoming_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Bersih()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        FormStencilInventory.Close()
    End Sub
End Class