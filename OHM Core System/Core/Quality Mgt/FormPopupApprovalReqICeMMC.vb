Imports MySql.Data.MySqlClient
Public Class FormPopupApprovalReqICeMMC
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Call connection()
            Dim str As String
            str = "UPDATE prodsys2_req_ic_emmc_tbl set is_active = '1' WHERE id = '" & TextBox9.Text & "'"
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Dispose()
            conn.Close()
            Me.Close()
        End Try
    End Sub
End Class