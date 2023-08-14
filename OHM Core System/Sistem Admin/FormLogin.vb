Public Class FormLogin

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox1.Focus()
    End Sub
    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call connection()
        'TemplateDasboardAdmin_VBNET.FormMain.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("FORM CAN'T EMPTY!", vbCritical)
                Call bersih()
            Else
                str = "SELECT username, password FROM prodsys2_account_tbl WHERE username='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        FormMain.Enabled = True
                        Me.Close()
                        OHM_Core_System.FormMain.LabelLogo.Visible = False
                    Else
                        MsgBox("WRONG USERNAME OR PASSWORD!", vbCritical)
                        Call bersih()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Try
                If TextBox1.Text = "" Or TextBox2.Text = "" Then
                    MsgBox("FORM CAN'T EMPTY!", vbCritical)
                    Call bersih()
                Else
                    str = "SELECT username, password FROM prodsys2_account_tbl WHERE username='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    rd.Read()
                    Try
                        If rd.HasRows Then
                            FormMain.Enabled = True
                            Me.Close()
                            FormMain.Show()
                            FormMain.LabelLogo.Visible = False
                        Else
                            MsgBox("WRONG USERNAME OR PASSWORD!", vbCritical)
                            Call bersih()
                        End If
                    Catch ex As Exception
                        MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                rd.Close()
            End Try
        End If
    End Sub
End Class