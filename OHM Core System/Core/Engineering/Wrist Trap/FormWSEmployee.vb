Public Class FormWSEmployee
    Private Sub FormWSEmployee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            FormWSChecker.TextBox1.Text = TextBox1.Text

            FormWSChecker.ShowDialog()
            Me.Close()
        End If
    End Sub
End Class