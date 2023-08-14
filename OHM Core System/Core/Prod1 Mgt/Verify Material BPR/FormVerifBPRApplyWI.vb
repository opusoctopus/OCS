
Public Class FormVerifBPRApplyWI
    Private Sub FormVerifBPRApplyWI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            FormWIBPR.ShowDialog()
        End If
    End Sub
End Class