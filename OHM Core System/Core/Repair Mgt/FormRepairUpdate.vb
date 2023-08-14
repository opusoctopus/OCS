Imports MySql.Data.MySqlClient
Imports System.IO
Public Class FormRepairUpdate

    Dim Myconnection As String = "server=10.217.4.115;user id=ohmuser;password=;database=stockflow_system" 'con2
    Dim result As Integer
    Dim imgpath As String
    Dim arrImage1() As Byte
    Dim arrImage2() As Byte
    Dim sql As String
    Dim conn As MySqlConnection
    Dim cmd As MySqlCommand
    Dim da As MySqlDataAdapter

    Sub CekAkun()
        conn = New MySqlConnection
        conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
        conn.Open()
        str = "SELECT * FROM prodsys2_repair_management_tbl WHERE txn_id = '" & TextBox11.Text & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                ComboBox1.Text = rd("status")
                If ComboBox1.Text = "DONE" Then
                    CheckBox4.Enabled = True
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                ElseIf ComboBox1.Text = "PROGRESS" Then
                    CheckBox4.Enabled = False
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End If
            Else
                CheckBox4.Enabled = False
                rd.Close()
                conn.Close()
                conn.Dispose()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FormRepairUpdate_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call CekAkun()
        TextBox10.Text = FormMain.LabelNama.Text
        TextBox11.ReadOnly = True
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox10.ReadOnly = True
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False

        Try
            conn = New MySqlConnection
            conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
            conn.Open()
            sql = "SELECT txn_id, evidence_repair_before, evidence_repair_after FROM prodsys2_repair_management_tbl WHERE txn_id = @txn_id;"
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                .Parameters.Clear()
                .Parameters.AddWithValue("@txn_id", TextBox11.Text)
                .ExecuteNonQuery()
            End With
            Dim arrImage1(), arrImage2() As Byte
            Dim dt As New DataTable
            da = New MySqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                arrImage1 = dt.Rows(0).Item(1)
                arrImage2 = dt.Rows(0).Item(2)

                Dim mstream1 As New MemoryStream(arrImage1)
                Dim mstream2 As New MemoryStream(arrImage2)

                PictureBox1.Image = Image.FromStream(mstream1)
                PictureBox2.Image = Image.FromStream(mstream2)
                Label22.Visible = False
                Label23.Visible = False
                conn.Close()
                conn.Dispose()
            Else
                PictureBox1.Image = Nothing
                PictureBox2.Image = Nothing
                Label22.Visible = True
                Label23.Visible = True
                conn.Close()
                conn.Dispose()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            PictureBox1.Image = Nothing
            PictureBox2.Image = Nothing
            Label22.Visible = True
            Label23.Visible = True
            conn.Close()
            conn.Dispose()
        End Try
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        str = "SELECT * FROM prodsys2_repair_management_tbl WHERE txn_id='" & TextBox11.Text & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                TextBox12.Text = rd("id")
                If PictureBox1.Image Is Nothing Or PictureBox2.Image Is Nothing Then
                    Try
                        conn = New MySqlConnection
                        conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                        conn.Open()
                        sql = "UPDATE prodsys2_repair_management_tbl SET status=@status, cause=@cause, cause2=@cause2, repairman=@repairman, position=@position, component=@component, remark=@remark, smt_repair=@smt_repair, drm_repair=@drm_repair, dft_repair=@dft_repair, verified_status=@verified_status, updated=@updated WHERE id = '" & TextBox12.Text & "';"
                        cmd = New MySqlCommand
                        With cmd
                            .Connection = conn
                            .CommandText = sql
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@status", ComboBox1.Text)
                            .Parameters.AddWithValue("@cause", TextBox13.Text)
                            .Parameters.AddWithValue("@cause2", ComboBox5.Text)
                            .Parameters.AddWithValue("@repairman", TextBox10.Text)
                            .Parameters.AddWithValue("@position", TextBox7.Text)
                            .Parameters.AddWithValue("@component", TextBox8.Text)
                            .Parameters.AddWithValue("@remark", TextBox9.Text)

                            If CheckBox1.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "CLOSED")
                                .Parameters.AddWithValue("@drm_repair", "")
                                .Parameters.AddWithValue("@dft_repair", "")
                            ElseIf CheckBox2.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "")
                                .Parameters.AddWithValue("@drm_repair", "CLOSED")
                                .Parameters.AddWithValue("@dft_repair", "")
                            ElseIf CheckBox3.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "")
                                .Parameters.AddWithValue("@drm_repair", "")
                                .Parameters.AddWithValue("@dft_repair", "CLOSED")
                            End If

                            If CheckBox4.Checked = True Then
                                .Parameters.AddWithValue("@verified_status", "VERIFIED")
                            Else
                                .Parameters.AddWithValue("@verified_status", "OPEN")
                            End If
                            .Parameters.AddWithValue("@updated", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss"))
                            .ExecuteNonQuery()
                            Call LoadData()
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                            Me.Close()
                        End With
                    Catch ex As MySqlException
                        MsgBox(ex.Message)
                    Finally
                        conn.Dispose()
                        conn.Close()
                    End Try
                Else
                    Try
                        Dim mstream1 As New System.IO.MemoryStream()
                        Dim mstream2 As New System.IO.MemoryStream()
                        PictureBox1.Image.Save(mstream1, System.Drawing.Imaging.ImageFormat.Jpeg)
                        PictureBox2.Image.Save(mstream2, System.Drawing.Imaging.ImageFormat.Jpeg)
                        arrImage1 = mstream1.GetBuffer()
                        arrImage2 = mstream2.GetBuffer()
                        Dim FileSize1 As UInt32
                        Dim FileSize2 As UInt32
                        FileSize1 = mstream1.Length
                        FileSize2 = mstream2.Length
                        mstream1.Close()
                        mstream2.Close()
                        conn = New MySqlConnection
                        conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                        conn.Open()
                        sql = "UPDATE prodsys2_repair_management_tbl SET status=@status, cause=@cause, cause2=@cause2, repairman=@repairman, position=@position, component=@component, remark=@remark, smt_repair=@smt_repair, drm_repair=@drm_repair, dft_repair=@dft_repair, verified_status=@verified_status, evidence_repair_before=@evidence_repair_before, evidence_repair_after=@evidence_repair_after, updated=@updated WHERE id = '" & TextBox12.Text & "';"
                        cmd = New MySqlCommand
                        With cmd
                            .Connection = conn
                            .CommandText = sql
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@status", ComboBox1.Text)
                            .Parameters.AddWithValue("@cause", TextBox13.Text)
                            .Parameters.AddWithValue("@cause2", ComboBox5.Text)
                            .Parameters.AddWithValue("@repairman", TextBox10.Text)
                            .Parameters.AddWithValue("@position", TextBox7.Text)
                            .Parameters.AddWithValue("@component", TextBox8.Text)
                            .Parameters.AddWithValue("@remark", TextBox9.Text)

                            If CheckBox1.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "CLOSED")
                                .Parameters.AddWithValue("@drm_repair", "")
                                .Parameters.AddWithValue("@dft_repair", "")
                            ElseIf CheckBox2.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "")
                                .Parameters.AddWithValue("@drm_repair", "CLOSED")
                                .Parameters.AddWithValue("@dft_repair", "")
                            ElseIf CheckBox3.Checked = True Then
                                .Parameters.AddWithValue("@smt_repair", "")
                                .Parameters.AddWithValue("@drm_repair", "")
                                .Parameters.AddWithValue("@dft_repair", "CLOSED")
                            End If

                            If CheckBox4.Checked = True Then
                                .Parameters.AddWithValue("@verified_status", "VERIFIED")
                            Else
                                .Parameters.AddWithValue("@verified_status", "OPEN")
                            End If
                            .Parameters.AddWithValue("@evidence_repair_before", arrImage1)
                            .Parameters.AddWithValue("@evidence_repair_after", arrImage2)
                            .Parameters.AddWithValue("@updated", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss"))
                            .ExecuteNonQuery()
                            Call LoadData()
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                            Me.Close()
                        End With
                    Catch ex As MySqlException
                        MsgBox(ex.Message)
                    Finally
                        conn.Dispose()
                        conn.Close()
                    End Try
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        FormEvidenceRepairBefore.ShowDialog()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        FormEvidenceRepairAfter.ShowDialog()
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

    Private Sub TextBox13_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox13.MouseClick
        TextBox13.SelectAll()
    End Sub
End Class