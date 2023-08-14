Public Class FormStarterPack
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormMain.ShowDialog()
    End Sub

    Private Sub FormStarterPack_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
    End Sub
End Class