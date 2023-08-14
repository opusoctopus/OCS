Imports MySql.Data.MySqlClient

Public Class FormMSL3MGT

    Sub customdgv2()
        With DataGridView1 'Ganti dengan nama datagridview kalian
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .BorderStyle = BorderStyle.None
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 30
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(200, 5, 83)
            .DefaultCellStyle.Font = New System.Drawing.Font("Arial", 10.0!)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 24
            End With
        End With
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 80 'txn_id
        DataGridView1.Columns(1).Width = 80 'txn_id
        DataGridView1.Columns(2).Width = 150 'txn_id
        DataGridView1.Columns(3).Width = 140 'pn lg
        DataGridView1.Columns(4).Width = 120 'lot id
        DataGridView1.Columns(5).Width = 80 'qty
        DataGridView1.Columns(6).Width = 100 'wo
        DataGridView1.Columns(7).Width = 100 'pn
        DataGridView1.Columns(8).Width = 80 'lot qty
        DataGridView1.Columns(9).Width = 100 'line
        DataGridView1.Columns(10).Width = 80 'shift
        DataGridView1.Columns(11).Width = 150 'start time
        DataGridView1.Columns(12).Width = 120 'run time
        DataGridView1.Columns(13).Width = 120 'end time
        DataGridView1.Columns(14).Width = 120 'pic
        DataGridView1.Columns(15).Width = 120 'remark
        DataGridView1.Columns(16).Width = 150 'date return
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE return_date IS NULL ORDER BY txn_id ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
    End Sub

    Sub showcount()
        Label7.Text = DataGridView1.RowCount
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormMSL3Data.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormMSL3Transaction.ShowDialog()
    End Sub

    Private Sub FormMSL3MGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim columnbutton1 As New DataGridViewButtonColumn
        columnbutton1.Width = 80
        columnbutton1.Name = "columnbutton1"
        columnbutton1.HeaderText = "Disposal"
        columnbutton1.Text = "Disposal"
        columnbutton1.UseColumnTextForButtonValue = True
        columnbutton1.FlatStyle = FlatStyle.Popup
        columnbutton1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(0, columnbutton1)

        Dim columnbutton2 As New DataGridViewButtonColumn
        columnbutton2.Width = 80
        columnbutton2.Name = "columnbutton2"
        columnbutton2.HeaderText = "Return"
        columnbutton2.Text = "Return"
        columnbutton2.UseColumnTextForButtonValue = True
        columnbutton2.FlatStyle = FlatStyle.Popup
        columnbutton2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(1, columnbutton2)

        Call customdgv2()
        Call LoadData()
        Call showcount()
        TextBox1.Select()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Baris As Integer = 0 To DataGridView1.Rows.Count - 1
            Dim now As DateTime = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            Dim start As DateTime = DataGridView1.Rows(Baris).Cells(11).Value
            Dim durasi As TimeSpan = New TimeSpan

            durasi = now - start
            DataGridView1.Rows(Baris).Cells(12).Value = durasi
            'Label16.Text = durasi.ToString

            If DataGridView1.Rows(Baris).Cells(12).Value > "7.00:00:00" Then
                DataGridView1.Rows(Baris).Cells(12).Style.BackColor = Color.FromArgb(200, 5, 83)
                DataGridView1.Rows(Baris).Cells(12).Style.ForeColor = Color.White
            ElseIf DataGridView1.Rows(Baris).Cells(12).Value > "3.00:00:00" Then
                DataGridView1.Rows(Baris).Cells(12).Style.BackColor = Color.Yellow
                DataGridView1.Rows(Baris).Cells(12).Style.ForeColor = Color.Black
            Else
                DataGridView1.Rows(Baris).Cells(12).Style.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormMSL3ReturnData.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormMSL3Return.TextBox1.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' txn ID
        FormMSL3Return.TextBox2.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' PN lg
        FormMSL3Return.TextBox3.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' lot id
        FormMSL3Return.TextBox4.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' qty
        FormMSL3Return.ComboBox1.Text = Me.DataGridView1.SelectedCells(7).Value.ToString ' line
        If Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 1" Then
            FormMSL3Return.RadioButton1.Checked = True
        ElseIf Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 2" Then
            FormMSL3Return.RadioButton2.Checked = True
        Else
            FormMSL3Return.RadioButton1.Checked = False
            FormMSL3Return.RadioButton2.Checked = False
        End If
        FormMSL3Return.TextBox5.Text = Me.DataGridView1.SelectedCells(11).Value.ToString ' start time
        FormMSL3Return.TextBox6.Text = Me.DataGridView1.SelectedCells(12).Value.ToString ' run time
        FormMSL3Return.TextBox7.Text = Me.DataGridView1.SelectedCells(15).Value.ToString ' remark
        FormMSL3Return.ShowDialog()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not TextBox1.Text = "" Then
                Try
                    str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE txn_id='" & TextBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView1.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call BuatKolom()
            Else
                Call LoadData()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "Change Model" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 1" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 2" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 3" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 4" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 5" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf ComboBox1.SelectedItem = "SMT Line 6" Then
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE line='" & ComboBox1.Text & "' AND return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        Else
            Try
                str = "SELECT txn_id as TxnID, pn_lg as PartNumberLG, lot_id as LotID, qty as Qty, wo as WO, pn as PN, lot_qty as LotQty, line as Line, shift as Shift, start_time as StartTime, run_time as RunTime, end_time as EndTime, pic as PIC, remark as Remark, return_date as ReturnDate FROM prodsys2_msl3_transaction_log_tbl WHERE return_date IS NULL ORDER BY txn_id ASC"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Return" Then
            FormMSL3Return.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' txn ID
            FormMSL3Return.TextBox2.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' PN lg
            FormMSL3Return.TextBox3.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' lot id
            FormMSL3Return.TextBox4.Text = Me.DataGridView1.SelectedCells(5).Value.ToString ' qty
            FormMSL3Return.ComboBox1.Text = Me.DataGridView1.SelectedCells(9).Value.ToString ' line
            If Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 1" Then
                FormMSL3Return.RadioButton1.Checked = True
            ElseIf Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 2" Then
                FormMSL3Return.RadioButton2.Checked = True
            Else
                FormMSL3Return.RadioButton1.Checked = False
                FormMSL3Return.RadioButton2.Checked = False
            End If
            FormMSL3Return.TextBox5.Text = Me.DataGridView1.SelectedCells(11).Value.ToString ' start time
            FormMSL3Return.TextBox6.Text = Me.DataGridView1.SelectedCells(12).Value.ToString ' run time
            FormMSL3Return.TextBox7.Text = Me.DataGridView1.SelectedCells(15).Value.ToString ' remark
            FormMSL3Return.ShowDialog()
        ElseIf DataGridView1.Columns(e.ColumnIndex).HeaderText = "Disposal" Then
            FormMSL3Disposal.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' txn ID
            FormMSL3Disposal.TextBox2.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' PN lg
            FormMSL3Disposal.TextBox3.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' lot id
            FormMSL3Disposal.TextBox4.Text = Me.DataGridView1.SelectedCells(5).Value.ToString ' qty
            FormMSL3Disposal.ComboBox1.Text = Me.DataGridView1.SelectedCells(9).Value.ToString ' line
            If Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 1" Then
                FormMSL3Disposal.RadioButton1.Checked = True
            ElseIf Me.DataGridView1.SelectedCells(10).Value.ToString = "Shift 2" Then
                FormMSL3Disposal.RadioButton2.Checked = True
            Else
                FormMSL3Disposal.RadioButton1.Checked = False
                FormMSL3Disposal.RadioButton2.Checked = False
            End If
            FormMSL3Disposal.TextBox5.Text = Me.DataGridView1.SelectedCells(11).Value.ToString ' start time
            FormMSL3Disposal.TextBox6.Text = Me.DataGridView1.SelectedCells(12).Value.ToString ' run time
            FormMSL3Disposal.TextBox7.Text = Me.DataGridView1.SelectedCells(15).Value.ToString ' remark
            FormMSL3Disposal.ShowDialog()
        End If
    End Sub
End Class