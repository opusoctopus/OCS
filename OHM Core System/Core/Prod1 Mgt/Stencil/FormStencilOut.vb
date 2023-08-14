Imports MySql.Data.MySqlClient

Public Class FormStencilOut

    Sub Bersih()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox13.Select()
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
        FormStencilMGT.DataGridView2.Columns(0).Width = 130
        FormStencilMGT.DataGridView2.Columns(1).Width = 120
        FormStencilMGT.DataGridView2.Columns(2).Width = 80
        FormStencilMGT.DataGridView2.Columns(3).Width = 80
        FormStencilMGT.DataGridView2.Columns(4).Width = 100
        FormStencilMGT.DataGridView2.Columns(5).Width = 100
        FormStencilMGT.DataGridView2.Columns(6).Width = 100
        FormStencilMGT.DataGridView2.Columns(7).Width = 100
        FormStencilMGT.DataGridView2.Columns(8).Width = 100
        FormStencilMGT.DataGridView2.Columns(9).Width = 100
        FormStencilMGT.DataGridView2.Columns(10).Width = 130
    End Sub

    Sub BuatKolomInLine()
        FormStencilMGT.DataGridView3.Columns(0).Width = 80
        FormStencilMGT.DataGridView3.Columns(1).Width = 120
        FormStencilMGT.DataGridView3.Columns(2).Width = 120
        FormStencilMGT.DataGridView3.Columns(3).Width = 80
        FormStencilMGT.DataGridView3.Columns(4).Width = 130
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        str = "SELECT * FROM prodsys2_stencil_rack WHERE kode_stencil='" & TextBox13.Text & "' AND status ='RACK'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                rd.Close()
                Dim cmd As New MySqlCommand("UPDATE prodsys2_stencil_rack SET status = @stat WHERE kode_stencil = @kode_stencil", conn)
                cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                cmd.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "LINE"
                cmd.ExecuteNonQuery()

                'Call connection()
                cmd = New MySql.Data.MySqlClient.MySqlCommand
                cmd.Connection = conn
                str = "INSERT INTO `prodsys2_stencil_trans_out_tbl` (`kode_stencil`, `pn_pcb`, `layer`, `line`, `pic`,`status`, `created`)
VALUES (@kode_stencil,@pn_pcb,@layer,@line, @pic, @status, @created)"
                cmd.Parameters.Add("@kode_stencil", MySqlDbType.VarChar).Value = TextBox13.Text
                cmd.Parameters.Add("@pn_pcb", MySqlDbType.VarChar).Value = TextBox14.Text
                cmd.Parameters.Add("@layer", MySqlDbType.VarChar).Value = ComboBox1.Text
                cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox5.Text
                cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = "AVAILABLE"
                cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                cmd.CommandText = str
                Try
                    cmd.ExecuteNonQuery()
                    Call Bersih()
                    Call LoadDataInLine()
                    Call LoadDataInspection()

                    rd.Close()
                    conn.Dispose()
                    conn.Close()
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try
            Else
                MsgBox("KODE STENCIL [" & TextBox13.Text & "] TIDAK ADA DI RAK ATAU MASIH RUNNING PRODUKSI!", MsgBoxStyle.Exclamation)
                rd.Close()
                Call Bersih()
                TextBox13.Select()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub FormStencilOut_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Bersih()
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Call connection()
            cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT kode_stencil, pn_pcb, layer, usages FROM prodsys2_stencil_inventory_tbl WHERE kode_stencil='" & TextBox13.Text & "'", conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                If rd("usages") < 40000 Then
                    TextBox14.Text = rd("pn_pcb")
                    ComboBox1.Text = rd("layer")
                    rd.Close()
                ElseIf rd("usages") < 50000 Then
                    MsgBox("[SYSTEM-WARNING] - PEMAKAIAN METALMASK INI [" & rd("usages") & "] SUDAH HAMPIR HABIS!", MsgBoxStyle.Exclamation)
                    rd.Close()
                ElseIf rd("usages") >= 50000 Then
                    MsgBox("[SYSTEM-ALERT] - TANSAKSI DIBATALKAN, PEMAKAIAN METALMASK INI SUDAH MELEBIHI BATAS MAXIMUM PEMAKAIAN!", MsgBoxStyle.Critical)
                    rd.Close()
                    Call Bersih()
                End If
                rd.Close()
                conn.Close()
                conn.Dispose()
            Else
                MsgBox("KODE STENCIL [" & TextBox13.Text & "] TIDAK TERDAFTAR!", MsgBoxStyle.Exclamation)
                Call Bersih()
                rd.Close()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class