Public Class FormPopupFingerCoat

    Private Sub FormPopupFingerCoat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Static tickCount As Integer = 1

        If Not Label1.Visible OrElse tickCount = 4 Then
            Label1.Visible = Not Label1.Visible
            tickCount = 1
        Else
            tickCount += 1
        End If
    End Sub
End Class