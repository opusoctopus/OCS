Imports System.IO
Imports MySql.Data.MySqlClient

Public Class FormRepairInspectionDetail

    Private Sub FormRepairInspectionDetail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        End If
    End Sub
End Class