Imports System.Windows.Forms.DataVisualization.Charting
Imports MySql.Data.MySqlClient
Imports System
Imports System.IO.Ports

Public Class FormWristtrapMGT

    Dim sql As String
    Dim result As Integer
    Dim myPort As Array 'COM Ports detected on the system will be stored here
    Delegate Sub SetTextCallback(ByVal [text] As String) 'Added to prevent threading errors during receiveing of data
    'Private CountDown As Integer = 10

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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
        With DataGridView2 'Ganti dengan nama datagridview kalian
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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
        With DataGridView3 'Ganti dengan nama datagridview kalian
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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
        With DataGridView4 'Ganti dengan nama datagridview kalian
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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Sub BuatKolomAttendance()
        DataGridView1.Columns(0).Width = 200 ' employee id
        DataGridView1.Columns(1).Width = 150 '1
        'DataGridView1.Columns(2).Width = 150 '1
        DataGridView1.Columns(2).Width = 870 '2
    End Sub

    Sub BuatKolomRemain()
        DataGridView2.Columns(0).Width = 150 ' employee id
        DataGridView2.Columns(1).Width = 280 '1
    End Sub

    Sub BuatKolomTotalCheck()
        DataGridView3.Columns(0).Width = 150 ' employee id
        DataGridView3.Columns(1).Width = 280 ' qty
    End Sub

    Sub LoadDataTotalCheck()
        Try
            str = "SELECT process as Process, COUNT(employee_id) AS Qty FROM prodsys2_wristrap_checker_tbl WHERE DATE(created) LIKE CURDATE() AND result='OK' GROUP BY process"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomTotalCheck()
    End Sub

    Sub BuatKolomTotalRemain()
        DataGridView4.Columns(0).Width = 150 ' employee id
        DataGridView4.Columns(1).Width = 80 ' qty
        DataGridView4.Columns(2).Width = 200 ' Rate
    End Sub

    Sub LoadDataTotalRemain()
        Try
            str = "SELECT total.process as Process, (total.qty - cek.qty) AS Total, concat(round(((cek.qty * 100) / total.qty),0), '%') AS Rate FROM prodsys2_wristrap_datatotal_tbl AS total, prodsys2_wristrap_dataremain_tbl AS cek WHERE total.process = cek.process"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView4.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomTotalRemain()
    End Sub

    Sub LoadDataAttendance()
        Try
            str = "SELECT full_name as EmployeeName, process as Process, result as Result FROM prodsys2_wristrap_checker_tbl WHERE DATE(created) LIKE CURDATE() ORDER BY created ASC;"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomAttendance()
    End Sub

    Sub LoadDataRemain()
        Try
            str = "SELECT process as Process, COUNT(employee_id) AS Qty FROM prodsys2_employee_tbl GROUP BY process"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomRemain()
    End Sub

    Sub ShowCountAttendance()
        Label3.Text = DataGridView1.RowCount
    End Sub

    Sub ShowCountRemain()
        Label4.Text = DataGridView4.RowCount
    End Sub

    Sub ComOpen()
        comPort.PortName = "COM4" 'Set SerialPort1 to the selected COM port at startup
        comPort.BaudRate = cmbBaud.Text 'Set Baud rate to the selected value on

        'Other Serial Port Property
        comPort.Parity = IO.Ports.Parity.None
        comPort.StopBits = IO.Ports.StopBits.One
        comPort.DataBits = 8 'Open our serial port
        comPort.Open()

        Button6.Enabled = False 'Disable Connect button
        Button7.Enabled = True 'and Enable Disconnect button
    End Sub

    Sub ComClose()
        comPort.Close() 'Close our Serial Port

        Button6.Enabled = True
        Button7.Enabled = False
    End Sub

    Private Sub FormWristtrapMGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadDataAttendance()
        Call LoadDataRemain()
        Call LoadDataTotalCheck()
        Call LoadDataTotalRemain()
        Call ShowCountAttendance()
        Call ShowCountRemain()
        Timer2.Stop()
        TextBox1.Select()
        GroupBox1.Hide()
        Label11.Text = System.DateTime.Now.ToString("MMMM yyyy")
        Label12.Text = System.DateTime.Now.ToString("dd")

        Dim months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames
        For Each month As String In months
            If Not String.IsNullOrEmpty(month) Then
                ComboBox1.Items.Add(month)
            End If
        Next

        ComboBox1.SelectedItem = System.DateTime.Now.ToString("MMMM")

        If ComboBox1.SelectedItem IsNot Nothing Then
            ComboBox1.Enabled = False
            Button3.Text = "DESELECT"
        Else
            ComboBox1.Enabled = True
            Button3.Text = "SELECT"
        End If

        With Chart1
            Dim employee() As String = {"Dept 1", "Dept 2", "Dept 3"}
            .Series.Clear()
            For a As Integer = 0 To employee.Length - 1
                .Series.Add(employee(a))

            Next
            .Series(0).ChartType = SeriesChartType.Pie
            .Series(1).ChartType = SeriesChartType.Pie
            .Series(2).ChartType = SeriesChartType.Pie

            .Series(0).Points.AddXY("MATERIAL", 80)
            .Series(0).Points.AddXY("SMT", 80)
            .Series(0).Points.AddXY("AOI", 60)
            .Series(0).Points.AddXY("BPR", 70)
            .Series(0).Points.AddXY("DRM", 50)
            .Series(0).Points.AddXY("DFT", 30)
            .Series(0).Points.AddXY("DMS", 20)
            .Series(0).Points.AddXY("LQC", 90)
        End With

        'When our form loads, auto detect all serial ports in the system and populate the cmbPort Combo box.
        myPort = IO.Ports.SerialPort.GetPortNames() 'Get all com ports available
        cmbBaud.Items.Add(9600) 'Populate the cmbBaud Combo box to common baud rates used
        cmbBaud.Items.Add(19200)
        cmbBaud.Items.Add(38400)
        cmbBaud.Items.Add(57600)
        cmbBaud.Items.Add(115200)

        For i = 0 To UBound(myPort)
            cmbPort.Items.Add(myPort(i))
        Next
        cmbPort.Text = cmbPort.Items.Item(0) 'Set cmbPort text to the first COM port detected
        cmbBaud.Text = cmbBaud.Items.Item(0) 'Set cmbBaud text to the first Baud rate on the list

        Button7.Enabled = False 'Initially Disconnect Button is Disabled
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label7.Text = System.DateTime.Now.ToString("HH:mm:ss")
        TextBox1.Select()

        For Baris As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(Baris).Cells(2).Value = "NG" Then
                DataGridView1.Rows(Baris).Cells(2).Style.ForeColor = Color.FromArgb(200, 5, 83)
            End If
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "SELECT" Then
            ComboBox1.Enabled = False
            Button3.Text = "DESELECT"
        ElseIf Button3.Text = "DESELECT" Then
            ComboBox1.Enabled = True
            Button3.Text = "SELECT"
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Button6.Enabled = False Then
                Label17.Text = ""
                RichTextBox1.Clear()
                'READ DATA EMPLOYEE
                Try
                    conn.Open()
                    str = "SELECT prodsys2_employee_tbl.full_name, prodsys2_employee_tbl.process, prodsys2_employee_tbl.department
                        FROM prodsys2_employee_tbl
                        LEFT JOIN prodsys2_wristrap_checker_tbl
                        ON prodsys2_employee_tbl.employee_id = prodsys2_wristrap_checker_tbl.employee_id
                        WHERE prodsys2_employee_tbl.employee_id = '" & TextBox1.Text & "'"
                    cmd = New MySqlCommand
                    With cmd
                        .Connection = conn
                        .CommandText = str
                    End With
                    rd = cmd.ExecuteReader()
                    If rd.HasRows Then
                        Do While rd.Read = True
                            TextBox2.Text = rd("full_name")
                            TextBox3.Text = rd("department")
                            TextBox4.Text = rd("process")
                        Loop
                        Timer2.Start()
                        Timer3.Enabled = True
                    Else
                        MsgBox("ID belum terdaftar!", MsgBoxStyle.Information)
                        TextBox1.Clear()
                        TextBox1.Select()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ProgressBar1.Value = 0
                    Timer2.Stop()
                    Timer3.Stop()
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    RichTextBox1.Clear()
                    Label17.Text = ""
                    TextBox1.Select()
                Finally
                    conn.Close()
                End Try
                rd.Close()
                conn.Dispose()
                conn.Close()
            Else
                MsgBox("Port communication isn't connected", MsgBoxStyle.Critical)
                TextBox1.Clear()
            End If
        End If
    End Sub

    Private Sub comPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles comPort.DataReceived
        ReceivedText(comPort.ReadExisting()) 'Automatically called every time a data is received at the serialPort
    End Sub

    Private Sub ReceivedText(ByVal [text] As String)
        'compares the ID of the creating Thread to the ID of the calling Thread
        If Me.RichTextBox1.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf ReceivedText)
            Me.Invoke(x, New Object() {(text)})
        Else
            Me.RichTextBox1.Text &= [text]
        End If
    End Sub

    Private Sub cmbPort_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPort.SelectedIndexChanged
        If comPort.IsOpen = False Then
            comPort.PortName = cmbPort.Text 'pop a message box to user if he is changing ports
        Else 'without disconnecting first.
            MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub cmbBaud_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBaud.SelectedIndexChanged
        If comPort.IsOpen = False Then
            comPort.BaudRate = cmbBaud.Text 'pop a message box to user if he is changing baud rate
        Else 'without disconnecting first.
            MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If RichTextBox1.Text.Contains("1.9") Or RichTextBox1.Text.Contains("2.") Then
            Label17.Text = "OK"
            Label17.ForeColor = Color.Green
            Timer3.Enabled = False
            'SIMPAN DATA OK
            ' create a command to insert
            Dim cmd As New MySqlCommand("INSERT INTO `prodsys2_wristrap_checker_tbl`
            (`employee_id`, `full_name`, `process`, `department`, `result`, `created`) VALUES (@employee_id,@full_name,@process,@department,@result, @created)", conn)

            ' add Parameters to the command
            cmd.Parameters.Add("@employee_id", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@full_name", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@department", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = Label17.Text
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("yyyy/MM/dd")
            conn.Open()
            If cmd.ExecuteNonQuery() = 1 Then
                'MsgBox("Data berhasil disimpan!", MsgBoxStyle.Information)
                ProgressBar1.Value = 0
                Timer2.Stop()
                Call UpdateData()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                RichTextBox1.Clear()
                'Label17.Text = ""
                TextBox1.Select()
                Call LoadDataAttendance()
                Call LoadDataRemain()
                Call LoadDataTotalRemain()
                Call LoadDataTotalCheck()
                Call ShowCountAttendance()
                Call ShowCountRemain()
            Else
                MessageBox.Show("ERROR")
            End If
            conn.Close()
            conn.Dispose()
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        ProgressBar1.Increment(10)
        If ProgressBar1.Value = 50 Then
            ProgressBar1.Value = 0
            If Label17.Text = "OK" Then
                Timer3.Enabled = False

                'SIMPAN DATA OK
                ' create a command to insert
                Dim cmd As New MySqlCommand("INSERT INTO `prodsys2_wristrap_checker_tbl`
            (`employee_id`, `full_name`, `process`, `department`, `result`, `created`) VALUES (@employee_id,@full_name,@process,@department,@result, @created)", conn)

                ' add Parameters to the command
                cmd.Parameters.Add("@employee_id", MySqlDbType.VarChar).Value = TextBox1.Text
                cmd.Parameters.Add("@full_name", MySqlDbType.VarChar).Value = TextBox2.Text
                cmd.Parameters.Add("@department", MySqlDbType.VarChar).Value = TextBox3.Text
                cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = TextBox4.Text
                cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = Label17.Text
                cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("yyyy/MM/dd")
                conn.Open()
                If cmd.ExecuteNonQuery() = 1 Then
                    'MsgBox("Data berhasil disimpan!", MsgBoxStyle.Information)
                    ProgressBar1.Value = 0
                    Timer2.Stop()
                    Call UpdateData()
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    RichTextBox1.Clear()
                    'Label17.Text = ""
                    TextBox1.Select()
                    Call LoadDataAttendance()
                    Call LoadDataRemain()
                    Call LoadDataTotalRemain()
                    Call LoadDataTotalCheck()
                    Call ShowCountAttendance()
                    Call ShowCountRemain()
                Else
                    MessageBox.Show("ERROR")
                End If
                conn.Close()
            Else
                Label17.Text = "NG"
                Label17.ForeColor = Color.Red
                Timer3.Stop()
                Dim tanya
                tanya = MessageBox.Show("DO YOU WANT TO RETRY?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If tanya = vbYes Then
                    Label17.Text = ""
                    RichTextBox1.Clear()
                    Timer2.Start()
                    Timer3.Start()
                Else
                    'DELETE LOG
                    sql = "DELETE FROM prodsys2_wristrap_checker_tbl WHERE DATE(created) LIKE CURDATE() AND result='NG' AND employee_id=" & TextBox1.Text
                    deleteData(sql)
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    RichTextBox1.Clear()
                    Timer2.Stop()
                    Timer3.Stop()
                End If
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call ComOpen()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Call ComClose()
    End Sub

    Private Sub deleteData(sql As String)
        Try
            Call connection()
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                result = .ExecuteNonQuery()
            End With
            If result > 0 Then
                'MsgBox("Log telah dihapus, harap cek kembali!", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call UpdateData()
    End Sub

    Sub UpdateData()
        sql = "Delete FROM prodsys2_wristrap_datatotal_tbl"
        deleteData(sql)

        sql = "Delete FROM prodsys2_wristrap_dataremain_tbl"
        deleteData(sql)

        'UPDATE DATA DAILY TOTAL CHECK EMPLOYEE
        conn.Open()
        For Baris As Integer = DataGridView2.Rows.Count - 1 To 0 Step -1
            cmd = New MySqlCommand("INSERT INTO `prodsys2_wristrap_datatotal_tbl`(`process`, `qty`, `updated`) VALUES (@process, @qty, @updated)", conn)
            cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = DataGridView2.Rows(Baris).Cells(0).Value.ToString()
            cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = DataGridView2.Rows(Baris).Cells(1).Value.ToString()
            cmd.Parameters.Add("@updated", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            cmd.ExecuteNonQuery()
        Next

        'UPDATE DAILY TOTAL DATA EMPLOYEE
        For Baris As Integer = DataGridView3.Rows.Count - 1 To 0 Step -1
            cmd = New MySqlCommand("INSERT INTO `prodsys2_wristrap_dataremain_tbl`(`process`, `qty`, `updated`) VALUES (@process, @qty, @updated)", conn)
            cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = DataGridView3.Rows(Baris).Cells(0).Value.ToString()
            cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = DataGridView3.Rows(Baris).Cells(1).Value.ToString()
            cmd.Parameters.Add("@updated", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            cmd.ExecuteNonQuery()
        Next
        'MsgBox("All data updated!", MsgBoxStyle.Information)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormWristtrapLog.ShowDialog()
    End Sub
End Class