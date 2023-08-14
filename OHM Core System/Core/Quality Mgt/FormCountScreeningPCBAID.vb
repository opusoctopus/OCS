Public Class FormCountScreeningPCBAID
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If TextBox1.Text = "0" Or TextBox1.Text = "" Then
                MsgBox("WARNING: An unknown error occured while input data", vbExclamation)
            Else
                FormScreeningPCBAID.Label11.Text = TextBox1.Text
                Me.Close()
                FormScreeningPCBAID.Enabled = True
                FormScreeningPCBAID.TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub FormCountScreeningPCBAID_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormScreeningPCBAID.Enabled = False
    End Sub
End Class