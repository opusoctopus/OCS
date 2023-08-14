Imports System.IO
Imports MySql.Data.MySqlClient

Public Class FormNozzleDetailInspection

    Dim result As Integer
    Dim imgpath As String
    Dim arrImage() As Byte
    Dim sql As String

    Sub BuatKolom()
        FormNozzleMGT.DataGridView1.Columns(0).Width = 60
        FormNozzleMGT.DataGridView1.Columns(1).Width = 110
        FormNozzleMGT.DataGridView1.Columns(2).Width = 90
        FormNozzleMGT.DataGridView1.Columns(3).Width = 120
        FormNozzleMGT.DataGridView1.Columns(4).Width = 120
        FormNozzleMGT.DataGridView1.Columns(5).Width = 100
        FormNozzleMGT.DataGridView1.Columns(6).Width = 290
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID,
                    nozzle_type as NozzleType, 
                    status_inspection as Judgment,
                    remark as Remark,
                    pic as PIC,
                    approved as Approved,
                    created as Created
                   FROM prodsys2_nozzle_inspection_tbl ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormNozzleMGT.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolom()
    End Sub

    Private Sub FormNozzleDetailInspection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        TextBox1.ReadOnly = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = False
        TextBox4.ReadOnly = True

        Dim img() As Byte
        Dim Query As String
        Call connection()
        Query = "SELECT * FROM prodsys2_nozzle_inspection_tbl WHERE txn_id = '" & Label6.Text & "'"
        cmd = New MySqlCommand(Query, conn)
        rd = cmd.ExecuteReader
        While rd.Read()
            img = rd.Item("evidence")
            Dim mstream As New MemoryStream(img)
            PictureBox1.Image = Image.FromStream(mstream)
        End While
        rd.Close()
        conn.Close()

        If FormMain.LabelNama.Text = "DOLLY MATHEUS" Or FormMain.LabelNama.Text = "KHOTIB" Or FormMain.LabelNama.Text = "SARJITO" Then
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            TextBox3.Enabled = True
            Button1.Enabled = True
        Else
            TextBox2.Enabled = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            TextBox3.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub saveData(sql As String)
        Try
            conn.Open()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckBox1.Checked = True Then
            sql = "UPDATE prodsys2_nozzle_inspection_tbl SET approved='" & CheckBox1.Text & "', approval='" & FormMain.LabelNama.Text & "', remark='" & TextBox3.Text & "' WHERE txn_id=" & Label6.Text
            saveData(sql)
            Call LoadData()
        ElseIf CheckBox2.Checked = True Then
            sql = "UPDATE prodsys2_nozzle_inspection_tbl SET approved='" & CheckBox2.Text & "', approval='" & FormMain.LabelNama.Text & "', remark='" & TextBox3.Text & "' WHERE txn_id=" & Label6.Text
            saveData(sql)
            Call LoadData()
        Else
            MsgBox("DATA GAGAL DI APPROVE!", MsgBoxStyle.Information)
        End If

        Me.Close()
    End Sub
End Class