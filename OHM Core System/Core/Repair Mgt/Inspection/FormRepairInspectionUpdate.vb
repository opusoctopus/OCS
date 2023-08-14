Imports System.IO
Imports MySql.Data.MySqlClient

Public Class FormRepairInspectionUpdate

    Dim result As Integer
    Dim imgpath As String
    Dim arrImage() As Byte
    Dim sql As String

    Sub BuatKolom()
        FormRepairInspection.DataGridView1.Columns(4).Width = 50 'TXN_ID '4
        FormRepairInspection.DataGridView1.Columns(5).Width = 80 'SN '5
        FormRepairInspection.DataGridView1.Columns(6).Width = 100 'WO '6
        FormRepairInspection.DataGridView1.Columns(7).Width = 100 'PN '7
        FormRepairInspection.DataGridView1.Columns(8).Width = 150 'NG '8
        FormRepairInspection.DataGridView1.Columns(9).Width = 60 'STATUS '9
        FormRepairInspection.DataGridView1.Columns(10).Width = 110 'REPAIRMAN '10
        FormRepairInspection.DataGridView1.Columns(11).Width = 80 'APPROVED '11
        FormRepairInspection.DataGridView1.Columns(12).Width = 120 'APPROVAL '12
        FormRepairInspection.DataGridView1.Columns(13).Width = 80 'APPROVED LQC '13
        FormRepairInspection.DataGridView1.Columns(14).Width = 120 'APPROVAL LQC '14
        FormRepairInspection.DataGridView1.Columns(15).Width = 100 'TIMESTAMPS '15
    End Sub

    Sub BuatKolomSUM()
        FormRepairInspection.DataGridView2.Columns(0).Width = 80 'SN
        FormRepairInspection.DataGridView2.Columns(1).Width = 100 'PROGRESS
        FormRepairInspection.DataGridView2.Columns(2).Width = 80 'APP LEADER
        FormRepairInspection.DataGridView2.Columns(3).Width = 80 'APP LQC
        FormRepairInspection.DataGridView2.Columns(4).Width = 60 'RESULT
    End Sub

    Sub showcount()
        FormRepairInspection.Label22.Text = FormRepairInspection.DataGridView1.RowCount
        FormRepairInspection.Label1.Text = FormRepairInspection.DataGridView2.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID, sn as BarCode, supply_wo as WO, pn as PartNumber, ng as NG, status as Status, repairman as Repairman, approved as VLeader, approval as Leader, approved_lqc as VLqc, approval_lqc as LQC, created as Timestamps FROM prodsys2_repair_inspection_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormRepairInspection.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()
    End Sub

    Sub LoadDataSum()
        Try
            str = "SELECT sn as SN, `status` AS `PROGRESS`, approved as LEADER, approved_lqc as LQC,
                       CASE
                           WHEN `status` = 'OPEN' AND approved = 'NO' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'NO' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'YES' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'NO' AND approved_lqc = 'YES' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'YES' AND approved_lqc = 'YES' THEN 'CLOSED'
                           ELSE 'Unknown'
                       END AS result
                   FROM prodsys2_repair_inspection_tbl;"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormRepairInspection.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomSUM()
    End Sub

    Private Sub FormRepairInspectionUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label9.Text = "PROGRESS" Or Label9.Text = "OPEN" Then
            Dim img() As Byte
            Dim Query As String
            Call connection()
            Query = "SELECT before_inspect FROM prodsys2_repair_inspection_tbl WHERE txn_id = '" & Label8.Text & "'"
            cmd = New MySqlCommand(Query, conn)
            rd = cmd.ExecuteReader
            While rd.Read()
                img = rd.Item("before_inspect")
                Dim mstream As New MemoryStream(img)
                PictureBox1.Image = Image.FromStream(mstream)
            End While
            rd.Close()
            conn.Close()

            PictureBox2.Image = Nothing

            ComboBox1.Enabled = True
            Button4.Enabled = True
        ElseIf Label9.Text = "DONE" Then
            Dim img() As Byte
            Dim img2() As Byte
            Dim Query As String
            Call connection()
            Query = "SELECT before_inspect, after_inspect FROM prodsys2_repair_inspection_tbl WHERE txn_id = '" & Label8.Text & "'"
            cmd = New MySqlCommand(Query, conn)
            rd = cmd.ExecuteReader
            While rd.Read()
                img = rd.Item("before_inspect")
                img2 = rd.Item("after_inspect")
                Dim mstream As New MemoryStream(img)
                Dim mstream2 As New MemoryStream(img2)
                PictureBox1.Image = Image.FromStream(mstream)
                PictureBox2.Image = Image.FromStream(mstream2)
            End While
            rd.Close()
            conn.Close()

            ComboBox1.Enabled = False
            Button4.Enabled = False
        ElseIf Label9.Text = "" Then
            MsgBox("DATA TIDAK DAPAT DIPROSES UPDATE!", MsgBoxStyle.Information)
            Me.Close()
        End If
    End Sub

    Private Sub saveData(sql As String)
        Try
            'conn.Open()
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                result = .ExecuteNonQuery()
            End With
            If result > 0 Then
                MsgBox("DATA SUDAH DIUPDATE!", MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.SelectedItem = "PROGRESS" Then
            sql = "UPDATE prodsys2_repair_inspection_tbl SET status='" & ComboBox1.Text & "' WHERE txn_id=" & Label8.Text
            saveData(sql)
            Call LoadData()
            Call LoadDataSum()
        ElseIf ComboBox1.SelectedItem = "DONE" Then
            Try
                sql = "UPDATE prodsys2_repair_inspection_tbl SET status=@status, after_inspect = @after_inspect, created = @created WHERE txn_id=" & Label8.Text
                Dim sqlcom As New MySqlCommand(sql, conn)
                Dim ms As New IO.MemoryStream()
                PictureBox2.Image.Save(ms, Imaging.ImageFormat.Jpeg)
                Dim data As Byte() = ms.GetBuffer()
                Dim p As New MySqlParameter("@after_inspect", MySqlDbType.LongBlob)
                p.Value = data
                sqlcom.Parameters.Add(p)
                sqlcom.Parameters.AddWithValue("@status", "DONE")
                sqlcom.Parameters.AddWithValue("@created", System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                conn.Open()
                sqlcom.ExecuteNonQuery()
                MsgBox("Data berhasil diupdate!", MsgBoxStyle.Information)
                Call LoadData()
                Call LoadDataSum()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        Else
            MsgBox("DATA GAGAL DI UPDATE!", MsgBoxStyle.Information)
        End If

        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        FormEvidenceRepairAfter.ShowDialog()
    End Sub
End Class