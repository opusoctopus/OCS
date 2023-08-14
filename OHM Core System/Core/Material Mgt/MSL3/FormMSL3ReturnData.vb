Imports MySql.Data.MySqlClient

Public Class FormMSL3ReturnData

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
            sql = "SELECT txn_id, pn_lg, lot_id, line, run_time, pic, remark, status, return_date FROM prodsys2_msl3_transaction_log_tbl WHERE end_time IS NOT NULL ORDER BY return_date DESC"
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = Sql
                dr = .ExecuteReader()
            End With
            ListView1.Items.Clear()
            Do While dr.Read = True
                Dim list = ListView1.Items.Add(dr(0)) 'txn id
                list.SubItems.Add(dr(1)) 'pn lg
                list.SubItems.Add(dr(2)) 'lot id
                list.SubItems.Add(dr(3)) 'line
                list.SubItems.Add(dr(4)) 'run time
                list.SubItems.Add(dr(5)) 'pic
                Dim objitem As Object = dr.GetValue(6) 'remark
                list.subitems.add(If(objitem IsNot Nothing, objitem.ToString, ""))
                Dim objitem2 As Object = dr.GetValue(7) 'status
                list.subitems.add(If(objitem2 IsNot Nothing, objitem2.ToString, ""))
                list.SubItems.Add(dr(8)) 'return date
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

    Private Sub FormMSL3ReturnData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call LoadData()
        Call showcount()
        'Call CustomListview()
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Try
                Call connection()
                sql = "SELECT txn_id, pn_lg, lot_id, line, run_time, pic, remark, status, return_date FROM prodsys2_msl3_transaction_log_tbl WHERE pn_lg='" & TextBox3.Text & "' ORDER BY return_date DESC"
                cmd = New MySqlCommand
                With cmd
                    .Connection = conn
                    .CommandText = sql
                    dr = .ExecuteReader()
                End With
                ListView1.Items.Clear()
                Do While dr.Read = True
                    Dim list = ListView1.Items.Add(dr(0))
                    list.SubItems.Add(dr(1))
                    list.SubItems.Add(dr(2))
                    list.SubItems.Add(dr(3))
                    list.SubItems.Add(dr(4))
                    list.SubItems.Add(dr(5))
                    Dim objitem As Object = dr.GetValue(6)
                    list.subitems.add(If(objitem IsNot Nothing, objitem.ToString, ""))
                    Dim objitem2 As Object = dr.GetValue(7) 'status
                    list.subitems.add(If(objitem2 IsNot Nothing, objitem2.ToString, ""))
                    list.SubItems.Add(dr(8))
                Loop
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            Finally
                conn.Close()
            End Try
        End If
    End Sub
End Class