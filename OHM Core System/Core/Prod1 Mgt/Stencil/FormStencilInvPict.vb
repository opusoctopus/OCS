Imports System.IO
Imports MySql.Data.MySqlClient

Public Class FormStencilInvPict
    Private Sub FormStencilInvPict_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim img() As Byte
        Dim Query As String
        Call connection()
        Query = "SELECT evidence1 FROM prodsys2_stencil_audit_data_tbl WHERE kode_stencil = '" & TextBox1.Text & "'"
        cmd = New MySqlCommand(Query, conn)
        rd = cmd.ExecuteReader
        Try
            While rd.Read()
                img = rd.Item("evidence1")
                Dim mstream As New MemoryStream(img)
                PictureBox1.Image = Image.FromStream(mstream)
                Label4.Hide()
            End While
            rd.Close()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label4.Show()
            rd.Close()
            conn.Close()
        End Try

    End Sub
End Class