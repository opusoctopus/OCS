Imports MySql.Data.MySqlClient
Public Class FormRepairInput

    Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox12.Text = ""
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        ComboBox3.SelectedItem = Nothing
        ComboBox4.SelectedItem = Nothing
        ComboBox6.SelectedItem = Nothing
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        TextBox5.Focus()
    End Sub

    Sub input()
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox12.Text = ""
        ComboBox6.SelectedItem = Nothing
        TextBox5.Focus()
    End Sub

    Sub BuatKolom()
        FormMasterRepair.DataGridView1.Columns(0).Width = 120 'TXN ID
        FormMasterRepair.DataGridView1.Columns(1).Width = 120 'STATUS
        FormMasterRepair.DataGridView1.Columns(2).Width = 120 'ORG
        FormMasterRepair.DataGridView1.Columns(3).Width = 120 'WO
        FormMasterRepair.DataGridView1.Columns(4).Width = 120 'PN
        FormMasterRepair.DataGridView1.Columns(5).Width = 180 'MODEL
        FormMasterRepair.DataGridView1.Columns(6).Width = 80 'LINE
        FormMasterRepair.DataGridView1.Columns(7).Width = 80 'TEAM
        FormMasterRepair.DataGridView1.Columns(8).Width = 60 'QTY
        FormMasterRepair.DataGridView1.Columns(9).Width = 80 'PROD
        FormMasterRepair.DataGridView1.Columns(10).Width = 100 'SN
        FormMasterRepair.DataGridView1.Columns(11).Width = 180 'CAUSE
        FormMasterRepair.DataGridView1.Columns(12).Width = 180 'CAUSE2
        FormMasterRepair.DataGridView1.Columns(13).Width = 120 'REQUESTOR
        FormMasterRepair.DataGridView1.Columns(14).Width = 120 'REPAIRMAN
        FormMasterRepair.DataGridView1.Columns(15).Width = 120 'POSITION
        FormMasterRepair.DataGridView1.Columns(16).Width = 120 'COMPONENT
        FormMasterRepair.DataGridView1.Columns(17).Width = 120 'REMARK
        FormMasterRepair.DataGridView1.Columns(18).Width = 120 'SMT REPAIR
        FormMasterRepair.DataGridView1.Columns(19).Width = 120 'DRM REPAIR
        FormMasterRepair.DataGridView1.Columns(20).Width = 120 'DFT REPAIR
        FormMasterRepair.DataGridView1.Columns(21).Width = 120 'VERIFIED STATUS
        FormMasterRepair.DataGridView1.Columns(22).Width = 130 'CREATED
        FormMasterRepair.DataGridView1.Columns(23).Width = 130 'UPDATED
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TXNID, status as Status, org as ORG, wo as WO, pn as PartNumber, model as Model, line as Line, team as Team, qty as Qty, product as Product, sn as BarCode, 
cause as Cause, cause2 as Cause2, requestor as Requestor, repairman as Repairman, position as Position, component as Component, remark as Remark, smt_repair as SMT, drm_repair as DRM, dft_repair as DFT, verified_status as Verified, created as Created, updated as Updated
FROM prodsys2_repair_management_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormMasterRepair.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
    End Sub

    Sub NomorOtomatis()
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM prodsys2_repair_management_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_repair_management_tbl)", conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            UrutanKode = "RO/" + System.DateTime.Now.ToString("yyyyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 9) + 1
            UrutanKode = "RO/" + System.DateTime.Now.ToString("yyyyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox11.Text = UrutanKode
        rd.Close()
        conn.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormRepairWOList.ShowDialog()
    End Sub

    Private Sub FormRepairInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call Bersih()
        ComboBox1.Text = "PROGRESS"
        ComboBox1.Enabled = False
        Call NomorOtomatis()
        CheckBox4.Enabled = False
        TextBox4.Text = "1"
        TextBox5.Focus()
        TextBox11.ReadOnly = True
        TextBox6.Text = FormMain.LabelNama.Text
        TextBox6.ReadOnly = True
        TextBox10.ReadOnly = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox5.Focus()
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox5.Text.StartsWith("IO")) Then
                MsgBox("SN [" & TextBox5.Text & "] INVALID CHARACTER", MsgBoxStyle.Information)
            Else
                conn.Open()
                str = "SELECT * FROM prodsys_sn_tbl WHERE sn='" & TextBox5.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        TextBox1.Text = rd("supply_wo")
                        TextBox2.Text = rd("pn")
                        rd.Close()
                        conn.Close()
                        conn.Dispose()

                        conn.Open()
                        str = "SELECT ebt, assembly_partno FROM prodsys_prod_plan_tbl WHERE ebt='" & TextBox2.Text & "'"
                        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                        rd = cmd.ExecuteReader
                        rd.Read()
                        Try
                            If rd.HasRows Then
                                TextBox3.Text = rd("assembly_partno")
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                                TextBox12.Focus()
                            End If
                        Catch ex As Exception
                            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End Try
                    Else
                        MsgBox("SN [" & TextBox5.Text & "] NOT FOUND, please input manuall!", vbExclamation)
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
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_repair_management_tbl` (`txn_id`, `status`, `org`, `wo` , `pn`, `model`, `line`, `team`, `qty`, `product`, `sn`, `cause`, `cause2`, `requestor`, `repairman`, `position`, `component`, `remark`, `smt_repair`, `drm_repair`, `dft_repair`, `verified_status`, `created`)
VALUES (@txn_id,@status,@org,@wo,@pn,@model,@line,@team,@qty,@product,@sn,@cause,@cause2,@requestor,@repairman,@position,@component,@remark,@smt_repair,@drm_repair,@dft_repair,@verified_status,@created)"
        cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox11.Text
        cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@org", MySqlDbType.VarChar).Value = ComboBox2.Text
        cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@model", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox3.Text
        cmd.Parameters.Add("@team", MySqlDbType.VarChar).Value = ComboBox4.Text
        cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox4.Text

        If CheckBox1.Checked = True Then
            cmd.Parameters.Add("@product", MySqlDbType.VarChar).Value = CheckBox1.Text
        ElseIf CheckBox2.Checked = True Then
            cmd.Parameters.Add("@product", MySqlDbType.VarChar).Value = CheckBox2.Text
        ElseIf CheckBox3.Checked = True Then
            cmd.Parameters.Add("@product", MySqlDbType.VarChar).Value = CheckBox3.Text
        End If

        cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox5.Text
        cmd.Parameters.Add("@cause", MySqlDbType.VarChar).Value = TextBox12.Text
        cmd.Parameters.Add("@cause2", MySqlDbType.VarChar).Value = ComboBox6.Text
        cmd.Parameters.Add("@requestor", MySqlDbType.VarChar).Value = TextBox6.Text
        cmd.Parameters.Add("@repairman", MySqlDbType.VarChar).Value = TextBox10.Text
        cmd.Parameters.Add("@position", MySqlDbType.VarChar).Value = TextBox7.Text
        cmd.Parameters.Add("@component", MySqlDbType.VarChar).Value = TextBox8.Text
        cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox9.Text

        If CheckBox1.Checked = True Then ' SMT
            cmd.Parameters.Add("@smt_repair", MySqlDbType.VarChar).Value = "OPEN"
        ElseIf CheckBox2.Checked = True Then ' BPR
            cmd.Parameters.Add("@drm_repair", MySqlDbType.VarChar).Value = "OPEN"
        ElseIf CheckBox3.Checked = True Then ' DFT
            cmd.Parameters.Add("@dft_repair", MySqlDbType.VarChar).Value = "OPEN"
        End If

        cmd.Parameters.Add("@verified_status", MySqlDbType.VarChar).Value = "OPEN"
        cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            Call BuatKolom()
            Call LoadData()
            Call input()
            Call NomorOtomatis()
            rd.Close()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT child, designators FROM prodsys2_master_bom_tbl WHERE designators='" & TextBox7.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    TextBox8.Text = rd("child")
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                    TextBox9.Focus()
                Else
                    MsgBox("Loc [" & TextBox7.Text & "] NOT FOUND, please check BOM manually!", vbExclamation)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                    TextBox9.Focus()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT desc_code FROM prodsys2_repair_code_master_tbl WHERE code='" & TextBox12.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    TextBox12.Text = ""
                    TextBox12.Text = rd("desc_code")
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                Else
                    MsgBox("Code [" & TextBox12.Text & "] NOT FOUND, please input manually!", vbExclamation)
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub

    Private Sub ComboBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT desc_code FROM prodsys2_repair_code_master_tbl WHERE code='" & ComboBox6.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    ComboBox6.Text = ""
                    ComboBox6.Text = rd("desc_code")
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                Else
                    MsgBox("Code [" & ComboBox6.Text & "] NOT FOUND, please input manually!", vbExclamation)
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox5.Focus()
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

    Private Sub TextBox4_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox4.MouseClick
        TextBox4.SelectAll()
    End Sub

    Private Sub TextBox5_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox5.MouseClick
        TextBox5.SelectAll()
    End Sub

    Private Sub TextBox6_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox6.MouseClick
        TextBox6.SelectAll()
    End Sub

    Private Sub TextBox7_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox7.MouseClick
        TextBox7.SelectAll()
    End Sub

    Private Sub TextBox8_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox8.MouseClick
        TextBox8.SelectAll()
    End Sub

    Private Sub TextBox9_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox9.MouseClick
        TextBox9.SelectAll()
    End Sub

    Private Sub TextBox10_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox10.MouseClick
        TextBox10.SelectAll()
    End Sub

    Private Sub TextBox11_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox11.MouseClick
        TextBox11.SelectAll()
    End Sub

    Private Sub TextBox12_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox12.MouseClick
        TextBox12.SelectAll()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        TextBox7.Focus()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        conn.Open()
        str = "SELECT child, designators FROM prodsys2_master_bom_tbl WHERE designators='" & TextBox7.Text & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                TextBox8.Text = rd("child")
                rd.Close()
                conn.Close()
                conn.Dispose()
                TextBox9.Focus()
            Else
                MsgBox("Loc [" & TextBox7.Text & "] NOT FOUND, please check BOM manually!", vbExclamation)
                rd.Close()
                conn.Close()
                conn.Dispose()
                TextBox9.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub
End Class