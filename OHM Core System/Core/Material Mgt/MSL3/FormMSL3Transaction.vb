Imports MySql.Data.MySqlClient

Public Class FormMSL3Transaction

    Dim cmd As MySqlCommand
    Dim dr As MySqlDataReader
    Dim sql As String
    Dim result As Integer

    Sub Bersih()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox5.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()

        ComboBox1.SelectedItem = Nothing

        RadioButton1.Checked = False
        RadioButton2.Checked = False

        TextBox2.Select()
    End Sub

    Sub NomorOtomatis()
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM prodsys2_msl3_transaction_log_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_msl3_transaction_log_tbl)", conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            UrutanKode = "MSL3" + System.DateTime.Now.ToString("yyyyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            UrutanKode = "MSL3" + System.DateTime.Now.ToString("yyyyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = UrutanKode
        rd.Close()
        conn.Close()
        conn.Dispose()
    End Sub

    Sub BuatKolom()
        FormMSL3MGT.DataGridView1.Columns(0).Width = 150 'txn_id
        FormMSL3MGT.DataGridView1.Columns(1).Width = 140 'pn lg
        FormMSL3MGT.DataGridView1.Columns(2).Width = 120 'lot id
        FormMSL3MGT.DataGridView1.Columns(3).Width = 80 'qty
        FormMSL3MGT.DataGridView1.Columns(4).Width = 100 'wo
        FormMSL3MGT.DataGridView1.Columns(5).Width = 100 'pn
        FormMSL3MGT.DataGridView1.Columns(6).Width = 80 'lot qty
        FormMSL3MGT.DataGridView1.Columns(7).Width = 100 'line
        FormMSL3MGT.DataGridView1.Columns(8).Width = 80 'shift
        FormMSL3MGT.DataGridView1.Columns(9).Width = 150 'start time
        FormMSL3MGT.DataGridView1.Columns(10).Width = 120 'run time
        FormMSL3MGT.DataGridView1.Columns(11).Width = 120 'end time
        FormMSL3MGT.DataGridView1.Columns(12).Width = 120 'pic
        FormMSL3MGT.DataGridView1.Columns(13).Width = 120 'remark
        FormMSL3MGT.DataGridView1.Columns(14).Width = 150 'date return
    End Sub

    Sub LoadData()
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
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.SelectedItem = Nothing Or RadioButton1.Checked = False AndAlso RadioButton2.Checked = False Then
            MsgBox("Data tidak lengkap!", vbExclamation)
            TextBox2.Select()
        Else
            If RadioButton1.Checked = True Then 'SHIFT 1
                conn.Open()
                cmd = New MySql.Data.MySqlClient.MySqlCommand
                cmd.Connection = conn
                str = "INSERT INTO `prodsys2_msl3_transaction_log_tbl`(`txn_id`, `pn_lg`, `lot_id`, `qty`, `wo`, `pn`, `lot_qty`, `line`, `shift`, `start_time`, `run_time`, `end_time`, `pic`, `remark`, `return_date`) 
VALUES (@txn_id,@pn_lg,@lot_id,@qty,@wo,@pn,@lot_qty,@line,@shift,@start_time,@run_time,@end_time,@pic,@remark,@return_date)"
                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                cmd.Parameters.Add("@pn_lg", MySqlDbType.VarChar).Value = TextBox2.Text
                cmd.Parameters.Add("@lot_id", MySqlDbType.VarChar).Value = TextBox3.Text
                cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox4.Text
                cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox8.Text
                cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox5.Text
                cmd.Parameters.Add("@lot_qty", MySqlDbType.VarChar).Value = TextBox9.Text
                cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox1.Text
                cmd.Parameters.Add("@shift", MySqlDbType.VarChar).Value = RadioButton1.Text
                cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
                cmd.Parameters.Add("@run_time", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.Parameters.Add("@end_time", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox10.Text
                cmd.Parameters.Add("@return_date", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.CommandText = str
                Try
                    cmd.ExecuteNonQuery()
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try

                FormMSL3PrintLabel.ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
                FormMSL3PrintLabel.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.FullPage
                FormMSL3PrintLabel.ShowDialog()

                Call Bersih()
                Call LoadData()
                Call showcount()
                Call NomorOtomatis()

            ElseIf RadioButton2.Checked = True Then 'SHIFT 2
                conn.Open()
                cmd = New MySql.Data.MySqlClient.MySqlCommand
                cmd.Connection = conn
                str = "INSERT INTO `prodsys2_msl3_transaction_log_tbl`(`txn_id`, `pn_lg`, `lot_id`, `qty`, `wo`, `pn`, `lot_qty`, `line`, `shift`, `start_time`, `run_time`, `end_time`, `pic`, `remark`, `return_date`) 
VALUES (@txn_id,@pn_lg,@lot_id,@qty,@wo,@pn,@lot_qty,@line,@shift,@start_time,@run_time,@end_time,@pic,@remark,@return_date)"
                cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
                cmd.Parameters.Add("@pn_lg", MySqlDbType.VarChar).Value = TextBox2.Text
                cmd.Parameters.Add("@lot_id", MySqlDbType.VarChar).Value = TextBox3.Text
                cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox4.Text
                cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox8.Text
                cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox5.Text
                cmd.Parameters.Add("@lot_qty", MySqlDbType.VarChar).Value = TextBox9.Text
                cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox1.Text
                cmd.Parameters.Add("@shift", MySqlDbType.VarChar).Value = RadioButton2.Text
                cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
                cmd.Parameters.Add("@run_time", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.Parameters.Add("@end_time", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox10.Text
                cmd.Parameters.Add("@return_date", MySqlDbType.VarChar).Value = DBNull.Value
                cmd.CommandText = str
                Try
                    cmd.ExecuteNonQuery()
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try

                FormMSL3PrintLabel.ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
                FormMSL3PrintLabel.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.FullPage
                FormMSL3PrintLabel.ShowDialog()

                Call Bersih()
                Call LoadData()
                Call showcount()
                Call NomorOtomatis()
            Else
                MsgBox("Shift belum dipilih!", vbInformation)
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FormMSL3Transaction_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        Call Bersih()
        Call NomorOtomatis()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox8.Focus()
        End If
    End Sub
End Class