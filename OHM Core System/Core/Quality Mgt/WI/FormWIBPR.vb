Imports System.IO
Imports MySql.Data.MySqlClient
Public Class FormWIBPR

    Private Sub FormWIBPR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim img() As Byte
        Dim Query As String
        Call connection()
        Query = "SELECT wi FROM prodsys2_wi_tbl WHERE ebt = '" & FormVerifBPRApplyWI.TextBox1.Text & "' "
        cmd = New MySql.Data.MySqlClient.MySqlCommand(Query, conn)
        rd = cmd.ExecuteReader
        While rd.Read()
            img = rd.Item("wi")
            Dim mstream As New MemoryStream(img)
            PictureBox1.Image = Image.FromStream(mstream)
            Label1.Hide()
        End While
        conn.Close()
        TextBox1.Text = ""
        TextBox1.Focus()
        Label2.Hide()
        Label3.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox1.Text.StartsWith("IO")) Then
                MsgBox("Invalid Format!", vbInformation)
                TextBox1.Text = ""
                TextBox1.Focus()
            Else
                conn.Open()
                str = "SELECT * FROM prodsys2_repair_management_tbl WHERE sn='" & TextBox1.Text & "' AND smt_repair = 'OPEN' AND verified_status = 'OPEN'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        Label3.Text = TextBox1.Text
                        Label2.Show()
                        Label3.Show()
                        TextBox1.Text = ""
                        TextBox1.Focus()
                        'My.Computer.Audio.Play("D:\ANDON-OHM\NG_OHM.wav")
                    Else
                        TextBox1.Text = ""
                        TextBox1.Focus()
                        Label2.Hide()
                        Label3.Hide()
                        'My.Computer.Audio.Play("D:\ANDON-OHM\OK_OHM.wav")
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
                TextBox1.Text = ""
                TextBox1.Focus()
            End If
        End If
    End Sub
End Class