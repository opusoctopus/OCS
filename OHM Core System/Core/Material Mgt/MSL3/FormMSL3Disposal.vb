Imports MySql.Data.MySqlClient

Public Class FormMSL3Disposal

    Sub Bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()

        TextBox2.Select()
    End Sub

    Sub BuatKolom()
        FormMSL3MGT.DataGridView1.Columns(0).Width = 80 'txn_id
        FormMSL3MGT.DataGridView1.Columns(1).Width = 80 'txn_id
        FormMSL3MGT.DataGridView1.Columns(2).Width = 150 'txn_id
        FormMSL3MGT.DataGridView1.Columns(3).Width = 140 'pn lg
        FormMSL3MGT.DataGridView1.Columns(4).Width = 120 'lot id
        FormMSL3MGT.DataGridView1.Columns(5).Width = 80 'qty
        FormMSL3MGT.DataGridView1.Columns(6).Width = 100 'wo
        FormMSL3MGT.DataGridView1.Columns(7).Width = 100 'pn
        FormMSL3MGT.DataGridView1.Columns(8).Width = 80 'lot qty
        FormMSL3MGT.DataGridView1.Columns(9).Width = 100 'line
        FormMSL3MGT.DataGridView1.Columns(10).Width = 80 'shift
        FormMSL3MGT.DataGridView1.Columns(11).Width = 150 'start time
        FormMSL3MGT.DataGridView1.Columns(12).Width = 120 'run time
        FormMSL3MGT.DataGridView1.Columns(13).Width = 120 'end time
        FormMSL3MGT.DataGridView1.Columns(14).Width = 120 'pic
        FormMSL3MGT.DataGridView1.Columns(15).Width = 120 'remark
        FormMSL3MGT.DataGridView1.Columns(16).Width = 150 'date return
    End Sub

    Sub LoadDataMSL3()
        Try
            str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE return_date IS NULL ORDER BY txn_id ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormMSL3MGT.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
    End Sub

    Sub showcount()
        FormMSL3MGT.Label7.Text = FormMSL3MGT.DataGridView1.RowCount
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For Baris As Integer = 0 To FormMSL3MGT.DataGridView1.Rows.Count - 1
            rd.Close()
            Call connection()
            str = "SELECT * FROM prodsys2_msl3_transaction_log_tbl WHERE txn_id='" & TextBox1.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    Dim sql As String
                    rd.Close()
                    sql = "UPDATE prodsys2_msl3_transaction_log_tbl SET line = @line, run_time = @run_time, end_time = @end_time, remark = @remark, status=@status, return_date = @return_date WHERE txn_id = @txn_id"
                    cmd.Connection = conn
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@txn_id", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@line", ComboBox1.Text)
                    cmd.Parameters.AddWithValue("@run_time", TextBox6.Text)
                    cmd.Parameters.AddWithValue("@end_time", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                    cmd.Parameters.AddWithValue("@remark", TextBox7.Text)
                    cmd.Parameters.AddWithValue("@status", "DISPOSAL")
                    cmd.Parameters.AddWithValue("@return_date", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                    Dim r As Integer
                    r = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Call LoadDataMSL3()
                    Call showcount()

                    Me.Close()
                Else

                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next
        rd.Close()
        conn.Close()
        conn.Dispose()
    End Sub

    Private Sub FormMSL3Disposal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class