Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports System
Imports System.ComponentModel
Imports System.Threading
Imports System.IO.Ports
Imports Microsoft.Reporting.WinForms
Imports System.IO
Imports System.Drawing.Imaging

Public Class FormCompletion
    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Delegate Sub SetTextCallback(ByVal [text] As String) 'Added to prevent threading errors during receiveing of data

    Sub Bersih()
        DataGridView1.Columns.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        'GroupBox1.Enabled = False
        Call showcount()
        Button3.Enabled = True
    End Sub

    Sub Cancel()
        rd.Close()
        conn.Close()
        conn.Dispose()
        TextBox1.Text = ""
        Label11.Text = ""
        Label19.Text = ""
        Label20.Text = ""
        Call NomorOtomatis()
        DataGridView1.Columns.Clear()
        Call BuatKolom()
        Call LoadData()
        Button3.Enabled = True
        Call showcount()
    End Sub

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
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Txn ID", "Txn ID") '0
        DataGridView1.Columns.Add("Type Box", "Type Box") '1
        DataGridView1.Columns.Add("WO", "WO") '2
        DataGridView1.Columns.Add("PN", "PN") '3
        DataGridView1.Columns.Add("Line", "Line") '4
        DataGridView1.Columns.Add("Shift", "Shift") '5
        DataGridView1.Columns.Add("SN", "SN") '6
        DataGridView1.Columns.Add("PIC", "PIC") '7
        DataGridView1.Columns.Add("Created", "Created") '8
        DataGridView1.Columns(0).Width = 150
        DataGridView1.Columns(1).Width = 90
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).Width = 60
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(8).Width = 165
    End Sub

    Sub BuatKolom2()
        'DataGridView2.Columns.Clear()
        DataGridView2.Columns(0).Width = 160
        DataGridView2.Columns(1).Width = 100
        DataGridView2.Columns(2).Width = 100
        DataGridView2.Columns(3).Width = 60
        DataGridView2.Columns(4).Width = 100
        DataGridView2.Columns(5).Width = 80
        DataGridView2.Columns(6).Width = 80
        DataGridView2.Columns(7).Width = 142
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID, pn as PN, type_box as TypeBOX, qty_box as Box, pic as PIC, remark as Remark, weight as Weight, printed_date as PrintedDate FROM prodsys2_labelbox_cskd_printed_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom2()
    End Sub

    Sub LoadDataPencarian()
        Try
            str = "SELECT txn_id, status, wo, org, pn, model, line, shift, qty, product, sn, 
cause, requestor, repairman, position, component, remark, smt_repair_status, bpr_repair_status, dft_repair_status, verified_status, created, updated 
FROM prodsys2_repair_management_tbl WHERE sn LIKE '" & TextBox1.Text & "'"
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

    Sub NomorOtomatis()
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM prodsys2_labelbox_cskd_printed_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_labelbox_cskd_printed_tbl)", conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            UrutanKode = System.DateTime.Now.ToString("yyyyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            UrutanKode = System.DateTime.Now.ToString("yyyyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = UrutanKode
        rd.Close()
        conn.Close()
    End Sub

    Public Sub StripNonAlphabetCharacters(ByVal input As TextBox)
        ' pattern matches any character that is NOT A-Z (allows upper and lower case alphabets)
        Dim rx As New Regex("[^0-9]")
        If (rx.IsMatch(input.Text)) Then
            Dim startPosition As Integer = input.SelectionStart - 1
            input.Text = rx.Replace(input.Text, "")
            input.SelectionStart = startPosition
        End If
    End Sub

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
        Label7.Text = DataGridView2.RowCount
        Label11.Text = DataGridView1.RowCount
    End Sub

    Private Sub FormCompletion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Bersih()
        Call customdgv2()
        Call LoadData()
        Call BuatKolom()
        'Call BuatKolom2()
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True

        'When our form loads, auto detect all serial ports in the system and populate the cmbPort Combo box.
        myPort = IO.Ports.SerialPort.GetPortNames() 'Get all com ports available
        cmbBaud.Items.Add(9600)     'Populate the cmbBaud Combo box to common baud rates used   
        cmbBaud.Items.Add(19200)
        cmbBaud.Items.Add(38400)
        cmbBaud.Items.Add(57600)
        cmbBaud.Items.Add(115200)

        For i = 0 To UBound(myPort)
            cmbPort.Items.Add(myPort(i))
        Next
        cmbPort.Text = cmbPort.Items.Item(0)    'Set cmbPort text to the first COM port detected
        cmbBaud.Text = cmbBaud.Items.Item(0)    'Set cmbBaud text to the first Baud rate on the list

        Button7.Enabled = False

        Call showcount()
        'Label15.Hide()
        Call NomorOtomatis()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GroupBox1.Enabled = True
        Call NomorOtomatis()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call Bersih()
        Call BuatKolom()
        Call NomorOtomatis()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormBoxCSKDWOList.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SerialPort1.PortName = cmbPort.Text         'Set SerialPort1 to the selected COM port at startup
        SerialPort1.BaudRate = cmbBaud.Text         'Set Baud rate to the selected value on 

        'Other Serial Port Property
        SerialPort1.Parity = IO.Ports.Parity.None
        SerialPort1.StopBits = IO.Ports.StopBits.One
        SerialPort1.DataBits = 8            'Open our serial port
        SerialPort1.Open()

        Button6.Enabled = False          'Disable Connect button
        Button7.Enabled = True       'and Enable Disconnect button

        GroupBox1.Enabled = True
        Call NomorOtomatis()
    End Sub

    Private Sub ReceivedText(ByVal [text] As String)
        'compares the ID of the creating Thread to the ID of the calling Thread
        If Me.TextBox4.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf ReceivedText)
            Me.Invoke(x, New Object() {(text)})
        Else
            Me.TextBox4.Text &= [text]
        End If
    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        ReceivedText(SerialPort1.ReadExisting())    'Automatically called every time a data is received at the serialPort
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        SerialPort1.Close()             'Close our Serial Port

        Button6.Enabled = True
        Button7.Enabled = False
        'Call Bersih()
        'Call BuatKolom()

        'GroupBox1.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_labelbox_cskd_printed_tbl`(`txn_id`, `pn`, `type_box`, `qty_box`, `pic`, `remark`, `weight`, `printed_date`) 
VALUES (@txn_id,@pn,@type_box,@qty_box,@pic,@remark,@weight,@printed_date)"
        cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@qty_box", MySqlDbType.VarChar).Value = Label22.Text
        cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
        cmd.Parameters.Add("@printed_date", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try

        conn.Close()
        For Each row As DataGridViewRow In DataGridView1.Rows
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_labelbox_cskd_log_tbl`(`txn_id`, `type_box`, `wo`, `pn` , `line`, `shift`, `sn`, `pic`, `created`) 
        VALUES (@txn_id,@type_box,@wo,@pn,@line,@shift,@sn,@pic,@created)"
            cmd.Parameters.AddWithValue("txn_id", row.Cells("Txn ID").Value)
            cmd.Parameters.AddWithValue("type_box", row.Cells("Type Box").Value)
            cmd.Parameters.AddWithValue("wo", row.Cells("WO").Value)
            cmd.Parameters.AddWithValue("pn", row.Cells("PN").Value)
            cmd.Parameters.AddWithValue("line", row.Cells("Line").Value)
            cmd.Parameters.AddWithValue("shift", row.Cells("Shift").Value)
            cmd.Parameters.AddWithValue("sn", row.Cells("SN").Value)
            cmd.Parameters.AddWithValue("pic", row.Cells("PIC").Value)
            cmd.Parameters.AddWithValue("created", row.Cells("Created").Value)
            cmd.CommandText = str
            Try
                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        Next

        Dim str2 As String = "SELECT txn_id, type_box, qty_box, pic, remark FROM prodsys2_labelbox_cskd_log_tbl WHERE txn_id='" & TextBox1.Text & "'"
        Try
            FormPrintLabelBoxCSKD.LabelBoxCSKD.Clear()
            FormPrintLabelBoxCSKD.LabelBoxCSKD.EnforceConstraints = False
            connection()
            da = New MySqlDataAdapter(str2, conn)
            da.Fill(FormPrintLabelBoxCSKD.LabelBoxCSKD.prodsys2_labelbox_cskd_log_tbl)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        FormPrintLabelBoxCSKD.ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
        FormPrintLabelBoxCSKD.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent
        FormPrintLabelBoxCSKD.ReportViewer1.ZoomPercent = 25
        FormPrintLabelBoxCSKD.ReportViewer1.RefreshReport()
        FormPrintLabelBoxCSKD.ShowDialog()

        rd.Close()
        conn.Close()
        conn.Dispose()
        TextBox1.Text = ""
        Label11.Text = ""
        Label19.Text = ""
        Label20.Text = ""
        Call NomorOtomatis()
        DataGridView1.Columns.Clear()
        Call BuatKolom()
        Call LoadData()
    End Sub

    Sub readingdata()
        Try
            If Not (TextBox4.Text.StartsWith("IO")) Then
                TextBox4.Text = ""
            Else
                conn.Open()
                str = "SELECT * FROM prodsys2_labelbox_cskd_log_tbl WHERE txn_id='" & TextBox1.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        Dim tanya
                        tanya = MessageBox.Show("TxnID sudah ada, replace ??", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If tanya = vbYes Then
                            rd.Close()
                            conn.Close()
                            conn.Dispose()

                            conn.Open()
                            cmd = New MySql.Data.MySqlClient.MySqlCommand
                            cmd.Connection = conn
                            str = "DELETE FROM prodsys2_labelbox_cskd_log_tbl WHERE txn_id='" & TextBox1.Text & "'"
                            cmd.CommandText = str
                            Try
                                cmd.ExecuteNonQuery()
                            Catch ex As Exception
                                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            End Try

                            Dim duplicate As Boolean = False
                            For Each row As DataGridViewRow In DataGridView1.Rows
                                For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                    If row.Cells(i).Value.ToString = TextBox4.Text Then
                                        duplicate = True
                                        row.Cells(i).Style.BackColor = Color.FromArgb(200, 5, 83)
                                        row.Cells(i).Style.ForeColor = Color.White
                                    End If
                                Next
                            Next

                            If (duplicate) Then
                                'MsgBox("DUPLICATE SERIAL NUMBER [" & TextBox4.Text & "]", vbCritical, ProductName)
                                TextBox4.Text = ""
                            Else
                                DataGridView1.Rows.Add(New String() {TextBox1.Text, TextBox2.Text, TextBox3.Text, ComboBox1.Text, ComboBox2.Text, TextBox4.Text, FormMain.LabelNama.Text, System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                                TextBox4.Text = ""
                                Call showcount()
                                rd.Close()
                                conn.Close()
                                conn.Dispose()
                            End If

                        Else
                            ' SAVE LOG PRINTED
                        End If
                    Else
                        rd.Close()
                        conn.Close()
                        conn.Dispose()

                        Dim duplicate As Boolean = False
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                If row.Cells(i).Value.ToString = TextBox4.Text Then
                                    duplicate = True
                                    row.Cells(i).Style.BackColor = Color.FromArgb(200, 5, 83)
                                    row.Cells(i).Style.ForeColor = Color.White
                                End If
                            Next
                        Next

                        If (duplicate) Then
                            TextBox4.Text = ""
                        Else
                            DataGridView1.Rows.Add(New String() {TextBox1.Text, TextBox2.Text, TextBox3.Text, ComboBox1.Text, ComboBox2.Text, TextBox4.Text, FormMain.LabelNama.Text, System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                            TextBox4.Text = ""
                            Call showcount()
                            rd.Close()
                            conn.Close()
                            conn.Dispose()
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try

                rd.Close()
                conn.Close()
                conn.Dispose()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If Len(TextBox4.Text) > 8 Then
            Try
                Call readingdata()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub cmbPort_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPort.SelectedIndexChanged
        If SerialPort1.IsOpen = False Then
            SerialPort1.PortName = cmbPort.Text         'pop a message box to user if he is changing ports
        Else                                            'without disconnecting first.
            'MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub cmbBaud_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBaud.SelectedIndexChanged
        If SerialPort1.IsOpen = False Then
            SerialPort1.BaudRate = cmbBaud.Text         'pop a message box to user if he is changing baud rate
        Else                                            'without disconnecting first.
            'MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call Cancel()
    End Sub
End Class