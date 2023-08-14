Public Class FormSetReminder
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If MaskedTextBox1.Text = "" Or MaskedTextBox2.Text = "" Or MaskedTextBox3.Text = "" Or MaskedTextBox4.Text = "" Then
            MsgBox("Semua Jam harus di set!", vbInformation)
        Else
            FormVerificationBPR.Label18.Text = MaskedTextBox1.Text
            FormVerificationBPR.Label19.Text = MaskedTextBox2.Text
            FormVerificationBPR.Label20.Text = MaskedTextBox3.Text
            FormVerificationBPR.Label21.Text = MaskedTextBox4.Text
        End If
        Me.Close()
    End Sub
End Class