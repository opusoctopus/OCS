Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions ' Regular Expressions Namespace
Public Class FormInventory
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox1.Focus()
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Assy", "Assy")
        DataGridView1.Columns.Add("Location", "Location")
        DataGridView1.Columns.Add("Serial Number", "Serial Number")
        DataGridView1.Columns.Add("Supply WO", "Supply WO")
        DataGridView1.Columns.Add("Part Number", "Part Number")
        DataGridView1.Columns.Add("Scanned by", "Scanned by")
        DataGridView1.Columns.Add("Created", "Created")
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).Width = 150
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).Width = 190
    End Sub

    Sub BuatKolom2()
        DataGridView2.Columns(0).Width = 100
        DataGridView2.Columns(1).Width = 90
        DataGridView2.Columns(2).Width = 100
        DataGridView2.Columns(3).Width = 100
        DataGridView2.Columns(4).Width = 100
        DataGridView2.Columns(5).Width = 130
    End Sub

    Sub BuatKolom3()
        DataGridView3.Columns(0).Width = 120
        DataGridView3.Columns(1).Width = 50
        DataGridView3.Columns(2).Width = 270
    End Sub

    Sub customdgv2()
        With DataGridView1 'Ganti dengan nama datagridview kalian
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 35
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
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 35
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
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            .AllowUserToAddRows = False
            .RowHeadersVisible = False
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 35
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

    Sub onn()
        GroupBox1.Enabled = True
        TextBox1.Focus()
    End Sub

    Sub off()
        GroupBox1.Enabled = False
    End Sub

    Sub LoadData()
        Try
            str = "select id_invassy as Inventory, location as Location, sn as SN, supply_wo as SupplyWO, pn as EBT, scanned_by as PIC, created as Created from prodsys2_inventory_assy_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom2()
        Try
            str = "select pn as EBT, count(sn) as qty, created as Created from prodsys2_inventory_assy_tbl group by pn"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        conn.Close()
        Call BuatKolom3()
    End Sub

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
        Label12.Text = DataGridView2.RowCount
        Label15.Text = DataGridView3.RowCount
    End Sub

    Private Sub FormInvSMT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call connection()
        Call off()
        Call customdgv2()
        Call LoadData()
        Call BuatKolom()
        Call BuatKolom2()
        Call BuatKolom3()
        With ComboBox1
            .Items.Add("BEFORE DFT")
            .Items.Add("BEFORE DMS")
            .Items.Add("AFTER DMS")
            .Items.Add("LOCATOR")
            .Items.Add("WAREHOUSE")
            .Items.Add("STOCK OVER")
        End With
        TextBox1.MaxLength = 10

        Label13.Text = ""
        Label11.Text = ""
        Call showcount()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call onn()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox1.Text.StartsWith("IO")) Then
                'My.Computer.Audio.Play("\\10.217.4.115\Shared Folder\SOFTWARE\PRODSYS OHM\ANDON-OHM\NG_OHM.wav")
                Label13.Text = "NG"
                Label13.Font = New Font("Arial", 48)
                Label13.ForeColor = Color.Red
                Label11.Text = "SN [" & TextBox1.Text & "] INVALID CHARACTER"
                Label11.Font = New Font("Arial Narrow", 14)
                Label11.ForeColor = Color.Crimson
                Call bersih()
            Else
                If ComboBox1.Text = "" Then
                    MsgBox("NO PROCESS IS SELECTED", vbInformation)
                    Call bersih()
                Else
                    conn.Open()
                    str = "SELECT * FROM prodsys_sn_tbl WHERE sn='" & TextBox1.Text & "'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    rd.Read()
                    Try
                        If rd.HasRows Then
                            Dim duplicate As Boolean = False
                            For Each row As DataGridViewRow In DataGridView1.Rows
                                For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                    If row.Cells(i).Value.ToString = TextBox1.Text Then
                                        duplicate = True
                                        row.Cells(i).Style.BackColor = Color.Crimson
                                        row.Cells(i).Style.ForeColor = Color.White
                                    End If
                                Next
                            Next

                            If (duplicate) Then
                                'My.Computer.Audio.Play("\\10.217.4.115\Shared Folder\SOFTWARE\PRODSYS OHM\ANDON-OHM\NG_OHM.wav")
                                Label13.Text = "NG"
                                Label13.Font = New Font("Arial", 48)
                                Label13.ForeColor = Color.Red
                                Label11.Text = "ALREADY SCANNED SN [" & TextBox1.Text & "]"
                                Label11.Font = New Font("Arial Narrow", 14)
                                Label11.ForeColor = Color.Crimson
                                Call bersih()
                                rd.Close()
                            Else
                                'My.Computer.Audio.Play("\\10.217.4.115\Shared Folder\SOFTWARE\PRODSYS OHM\ANDON-OHM\OK_OHM.wav")
                                TextBox1.Text = rd("sn")
                                TextBox2.Text = rd("supply_wo")
                                TextBox3.Text = rd("pn")
                                Label11.Text = ""
                                Label13.Text = "OK"
                                Label13.ForeColor = Color.Lime
                                Label13.Font = New Font("Arial", 48)
                                rd.Close()
                                conn.Close()

                                conn.Open()

                                DataGridView1.Rows.Add(New String() {ComboBox1.Text, TextBox5.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text, FormMain.LabelNama.Text, System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                                cmd = New MySql.Data.MySqlClient.MySqlCommand
                                cmd.Connection = conn
                                str = "INSERT INTO `prodsys2_inventory_assy_tbl`(`id_invassy`, `location`, `sn`, `supply_wo` , `pn`, `scanned_by`, `created`) VALUES (@id_invassy,@location,@sn,@supply_wo,@pn,@scanned_by,@created)"
                                cmd.Parameters.Add("@id_invassy", MySqlDbType.VarChar).Value = ComboBox1.Text
                                cmd.Parameters.Add("@location", MySqlDbType.VarChar).Value = TextBox5.Text
                                cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox1.Text
                                cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = TextBox2.Text
                                cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
                                cmd.Parameters.Add("@scanned_by", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                cmd.CommandText = str
                                Try
                                    cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    rd.Close()
                                    conn.Close()
                                    conn.Dispose()
                                End Try
                            End If
                            Call bersih()
                        Else
                            'Call connection()
                            rd.Close()

                            Dim duplicate As Boolean = False
                            For Each row As DataGridViewRow In DataGridView1.Rows
                                For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                    If row.Cells(i).Value.ToString = TextBox1.Text Then
                                        duplicate = True
                                        row.Cells(i).Style.BackColor = Color.Crimson
                                        row.Cells(i).Style.ForeColor = Color.White
                                    End If
                                Next
                            Next

                            If (duplicate) Then
                                'My.Computer.Audio.Play("\\10.217.4.115\Shared Folder\SOFTWARE\PRODSYS OHM\ANDON-OHM\NG_OHM.wav")
                                Label13.Text = "NG"
                                Label13.Font = New Font("Arial", 48)
                                Label13.ForeColor = Color.Red
                                Label11.Text = "ALREADY SCANNED SN [" & TextBox1.Text & "]"
                                Label11.Font = New Font("Arial Narrow", 14)
                                Label11.ForeColor = Color.Crimson
                                Call bersih()
                                rd.Close()
                            Else
                                'My.Computer.Audio.Play("\\10.217.4.115\Shared Folder\SOFTWARE\PRODSYS OHM\ANDON-OHM\OK_OHM.wav")
                                Label11.Text = ""
                                Label13.Text = "OK"
                                Label13.ForeColor = Color.Lime
                                Label13.Font = New Font("Arial", 48)

                                DataGridView1.Rows.Add(New String() {ComboBox1.Text, TextBox5.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text, FormMain.LabelNama.Text, System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                                cmd = New MySql.Data.MySqlClient.MySqlCommand
                                cmd.Connection = conn
                                str = "INSERT INTO `prodsys2_inventory_assy_tbl`(`id_invassy`, `location`, `sn`, `supply_wo` , `pn`, `scanned_by`, `created`) VALUES (@id_invassy,@location,@sn,@supply_wo,@pn,@scanned_by,@created)"
                                cmd.Parameters.Add("@id_invassy", MySqlDbType.VarChar).Value = ComboBox1.Text
                                cmd.Parameters.Add("@location", MySqlDbType.VarChar).Value = TextBox5.Text
                                cmd.Parameters.Add("@sn", MySqlDbType.VarChar).Value = TextBox1.Text
                                cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = TextBox2.Text
                                cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
                                cmd.Parameters.Add("@scanned_by", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
                                cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                                cmd.CommandText = str
                                Try
                                    cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    rd.Close()
                                    conn.Close()
                                    conn.Dispose()
                                End Try
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
                    Call bersih()
                    Call showcount()
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.Text IsNot Nothing Then
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call LoadData()
        Call BuatKolom2()
        Call BuatKolom3()
        Call showcount()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportDgvToExcel(Me.DataGridView2)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Call connection()
        Dim fromdate As String = DateTimePicker1.Value.ToString("dd/MM/yyyy")
        Dim todate As String = DateTimePicker2.Value.ToString("dd/MM/yyyy")

        Try
            str = "SELECT id_invassy,location, sn, supply_wo, pn, scanned_by, created FROM prodsys2_inventory_assy_tbl WHERE created BETWEEN'" & fromdate & "' AND '" & todate & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable

            da.Fill(dt)
            DataGridView2.DataSource = dt
            conn.Close()
            Call BuatKolom2()
            Call showcount()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
            Call BuatKolom2()
        End Try
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Try
                str = "SELECT id_invassy, location, sn, supply_wo, pn, scanned_by, created FROM prodsys2_inventory_assy_tbl WHERE pn LIKE '" & TextBox4.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable

                da.Fill(dt)
                DataGridView2.DataSource = dt
                conn.Close()
                Call BuatKolom2()
                Call showcount()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
                Call BuatKolom2()
            End Try
        End If
    End Sub
End Class