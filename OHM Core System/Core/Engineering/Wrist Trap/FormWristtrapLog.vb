Imports MySql.Data.MySqlClient

Public Class FormWristtrapLog

    Dim cmd As MySqlCommand
    Dim dr As MySqlDataReader
    Dim sql As String
    Dim result As Integer

    Sub CustomListview()
        Dim iView As Integer = ListView1.Items.Count - 1
        For i = 1 To iView Step 2
            ListView1.Items(i).BackColor = Drawing.Color.Crimson
            ListView1.Items(i).SubItems(1).BackColor = Drawing.Color.Crimson
            ListView1.Items(i).SubItems(2).BackColor = Drawing.Color.Crimson
        Next i
    End Sub

    Sub LoadData()
        Try
            Call connection()
            sql = "SELECT * FROM prodsys2_wristrap_checker_tbl WHERE employee_id IS NOT NULL ORDER BY created DESC"
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                dr = .ExecuteReader()
            End With
            ListView1.Items.Clear()
            Do While dr.Read = True
                Dim list = ListView1.Items.Add(dr(0)) 'txn id
                list.SubItems.Add(dr(1)) 'employee id
                list.SubItems.Add(dr(2)) 'employee name
                list.SubItems.Add(dr(3)) 'departmen
                list.SubItems.Add(dr(4)) 'result
                list.SubItems.Add(dr(5)) 'created
            Loop
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        Finally
            conn.Close()
        End Try
    End Sub

    Sub showcount()
        Label7.Text = ListView1.Items.Count
    End Sub

    Private Sub FormWristtrapLog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call LoadData()
        Call showcount()
        TextBox3.Select()
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If TextBox3.Text IsNot Nothing Then
                Try
                    Call connection()
                    sql = "SELECT * FROM prodsys2_wristrap_checker_tbl WHERE employee_id='" & TextBox3.Text & "'"
                    cmd = New MySqlCommand
                    With cmd
                        .Connection = conn
                        .CommandText = sql
                        dr = .ExecuteReader()
                    End With
                    ListView1.Items.Clear()
                    Do While dr.Read = True
                        Dim list = ListView1.Items.Add(dr(0)) 'txn id
                        list.SubItems.Add(dr(1)) 'employee id
                        list.SubItems.Add(dr(2)) 'employee name
                        list.SubItems.Add(dr(3)) 'departmen
                        list.SubItems.Add(dr(4)) 'result
                        list.SubItems.Add(dr(5)) 'created
                    Loop
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                Finally
                    conn.Close()
                    Call showcount()
                End Try
            ElseIf textbox3.Text = "" Then
                Call LoadData()
                Call showcount()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call LoadData()
        Call showcount()
        TextBox3.Clear()
        TextBox3.Select()
    End Sub
End Class