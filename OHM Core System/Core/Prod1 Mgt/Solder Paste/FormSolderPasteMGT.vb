Imports MySql.Data.MySqlClient

Public Class FormSolderPasteMGT

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
        With DataGridView5 'Ganti dengan nama datagridview kalian
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
        With DataGridView6 'Ganti dengan nama datagridview kalian
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
        With DataGridView7 'Ganti dengan nama datagridview kalian
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
        With DataGridView8 'Ganti dengan nama datagridview kalian
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
        With DataGridView9 'Ganti dengan nama datagridview kalian
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
        With DataGridView10 'Ganti dengan nama datagridview kalian
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

    Sub KolomCoolingCase()
        DataGridView1.Columns(0).Width = 150 'Barcode
        DataGridView1.Columns(1).Width = 120 'Maker
        DataGridView1.Columns(2).Width = 120 'Type name
        DataGridView1.Columns(3).Width = 120 'top/bott
        DataGridView1.Columns(4).Width = 180 'start time
        DataGridView1.Columns(5).Width = 180 'running time
        DataGridView1.Columns(6).Width = 180 'end time
        DataGridView1.Columns(7).Width = 120 'status
    End Sub

    Sub KolomCheckRight()
        DataGridView2.Columns(0).Width = 150 'Barcode
        DataGridView2.Columns(1).Width = 120 'Maker
        DataGridView2.Columns(2).Width = 120 'Type name
        DataGridView2.Columns(3).Width = 120 'top/bott
        DataGridView2.Columns(4).Width = 120 'hole
        DataGridView2.Columns(5).Width = 180 'start time
        DataGridView2.Columns(6).Width = 180 'running time
        DataGridView2.Columns(7).Width = 180 'end time
        DataGridView2.Columns(8).Width = 180 'status
        DataGridView2.Columns(9).Width = 180 'pic
    End Sub

    Sub KolomMixer()
        DataGridView3.Columns(0).Width = 150 'Barcode
        DataGridView3.Columns(1).Width = 120 'Maker
        DataGridView3.Columns(2).Width = 120 'Type name
        DataGridView3.Columns(3).Width = 120 'top/bott
        DataGridView3.Columns(4).Width = 180 'start time
        DataGridView3.Columns(5).Width = 180 'running time
        DataGridView3.Columns(6).Width = 180 'end time
        DataGridView3.Columns(7).Width = 180 'status
    End Sub

    Sub KolomLine()
        DataGridView4.Columns(0).Width = 150 'Barcode
        DataGridView4.Columns(1).Width = 120 'Maker
        DataGridView4.Columns(2).Width = 120 'Type name
        DataGridView4.Columns(3).Width = 120 'top/bott
        DataGridView4.Columns(4).Width = 120 'line
        DataGridView4.Columns(5).Width = 180 'start time
        DataGridView4.Columns(6).Width = 180 'running time
        DataGridView4.Columns(7).Width = 180 'end time
        DataGridView4.Columns(8).Width = 180 'status
        DataGridView4.Columns(9).Width = 180 'pic
    End Sub

    Sub KolomReadyMixer()
        DataGridView5.Columns(0).Width = 120 'txn id
        DataGridView5.Columns(1).Width = 80 'topbott
        DataGridView5.Columns(2).Width = 60 'hole
    End Sub

    Sub KolomInLine()
        DataGridView6.Columns(0).Width = 120 'txn id
        DataGridView6.Columns(1).Width = 80 'topbott
        DataGridView6.Columns(2).Width = 60 'line
    End Sub

    Sub KolomLogPemakaian()
        DataGridView7.Columns(0).Width = 180 'txn id
        DataGridView7.Columns(1).Width = 100 'topbott
        DataGridView7.Columns(2).Width = 180 'start
        DataGridView7.Columns(3).Width = 180 'end time
        DataGridView7.Columns(4).Width = 120 'run time
    End Sub

    Sub KolomLogPemakaianCheckright()
        DataGridView8.Columns(0).Width = 180 'txn id
        DataGridView8.Columns(1).Width = 100 'topbott
        DataGridView8.Columns(2).Width = 180 'start
        DataGridView8.Columns(3).Width = 180 'end
        DataGridView8.Columns(4).Width = 120 'run 
        DataGridView8.Columns(5).Width = 80 'pic
    End Sub

    Sub KolomLogPemakaianMixer()
        DataGridView9.Columns(0).Width = 180 'txn id
        DataGridView9.Columns(1).Width = 100 'topbott
        DataGridView9.Columns(2).Width = 180 'start
        DataGridView9.Columns(3).Width = 180 'end
        DataGridView9.Columns(4).Width = 120 'run 
    End Sub

    Sub KolomLogPemakaianLine()
        DataGridView10.Columns(0).Width = 180 'txn id
        DataGridView10.Columns(1).Width = 100 'topbott
        DataGridView10.Columns(2).Width = 180 'line
        DataGridView10.Columns(3).Width = 180 'start
        DataGridView10.Columns(4).Width = 120 'end 
        DataGridView10.Columns(5).Width = 80 'run
        DataGridView10.Columns(6).Width = 80 'pic
    End Sub

    Sub LoadDataCoolingCase()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE status = 'AVAILABLE' ORDER BY start_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCoolingCase()
    End Sub

    Sub LoadDataCheckRight()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, hole, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'AVAILABLE' ORDER BY start_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCheckRight()
    End Sub

    Sub LoadDataMixer()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE status = 'AVAILABLE' ORDER BY start_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomMixer()
    End Sub

    Sub LoadDataLine()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, line, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'AVAILABLE' ORDER BY start_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView4.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLine()
    End Sub

    Sub LoadDataReadyMixer()
        Try
            str = "SELECT txn_id as Barcode , topbott as Lot, hole as Hole FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView5.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomReadyMixer()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT txn_id as Barcode , topbott as Lot, line as Line FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView6.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomInLine()
    End Sub

    Sub LoadDataLogPemakaian()
        Try
            str = "SELECT txn_id, topbott, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, DATE_FORMAT(STR_TO_DATE(end_time,'%d/%m/%Y'), '%Y/%m/%d') as end_time, run_time FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE status = 'EMPTY' ORDER BY end_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView7.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLogPemakaian()
    End Sub

    Sub LoadDataPemakaianCheckright()
        Try
            str = "SELECT txn_id, topbott, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, DATE_FORMAT(STR_TO_DATE(end_time,'%d/%m/%Y'), '%Y/%m/%d') as end_time, run_time, pic FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'EMPTY' ORDER BY end_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView8.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLogPemakaianCheckright()
    End Sub

    Sub LoadDataPemakaianMixer()
        Try
            str = "SELECT txn_id, topbott, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, DATE_FORMAT(STR_TO_DATE(end_time,'%d/%m/%Y'), '%Y/%m/%d') as end_time, run_time FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE status = 'EMPTY' ORDER BY end_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView9.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLogPemakaianMixer()
    End Sub

    Sub LoadDataPemakaianLine()
        Try
            str = "SELECT txn_id, topbott, line, DATE_FORMAT(STR_TO_DATE(start_time,'%d/%m/%Y'), '%Y/%m/%d') as start_time, DATE_FORMAT(STR_TO_DATE(end_time,'%d/%m/%Y'), '%Y/%m/%d') as end_time, run_time, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'EMPTY' ORDER BY end_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView10.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLogPemakaianLine()
    End Sub

    Sub ShowCountCoolingCase()
        Label3.Text = DataGridView1.RowCount
        TabPage1.Text = "Cooling Case [" & DataGridView1.RowCount & "]"
    End Sub

    Sub ShowCountCheckRight()
        Label6.Text = DataGridView2.RowCount
        TabPage2.Text = "Check Right [" & DataGridView2.RowCount & "]"
    End Sub

    Sub ShowCountMixer()
        Label11.Text = DataGridView3.RowCount
        TabPage3.Text = "Mixer [" & DataGridView3.RowCount & "]"
    End Sub

    Sub ShowCountLine()
        Label7.Text = DataGridView4.RowCount
        TabPage4.Text = "Line [" & DataGridView4.RowCount & "]"
    End Sub

    Sub ShowCountReadyMixing()
        Label16.Text = DataGridView5.RowCount
    End Sub

    Sub ShowCountInline()
        Label17.Text = DataGridView6.RowCount
    End Sub

    Sub ShowCountLogPemakaian()
        TabPage5.Text = "History [" & DataGridView10.RowCount & "]"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormCoolingCase.ShowDialog()
    End Sub

    Private Sub FormSolderPasteMGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        'Call KolomCoolingCase()
        'Call KolomCheckRight()
        'Call KolomMixer()
        'Call KolomLine()

        Call LoadDataCoolingCase()
        Call LoadDataCheckRight()
        Call LoadDataMixer()
        Call LoadDataLine()
        Call LoadDataReadyMixer()
        Call LoadDataInLine()
        Call LoadDataLogPemakaian()
        Call LoadDataPemakaianCheckright()
        Call LoadDataPemakaianMixer()
        Call LoadDataPemakaianLine()

        Call ShowCountCoolingCase()
        Call ShowCountCheckRight()
        Call ShowCountMixer()
        Call ShowCountLine()
        Call ShowCountReadyMixing()
        Call ShowCountInline()
        Call ShowCountLogPemakaian()

        ComboBox1.SelectedItem = Nothing
        TextBox2.Select()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Baris As Integer = DataGridView1.Rows.Count - 1 To 0 Step -1
            Dim now As DateTime = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            Dim start As DateTime = DataGridView1.Rows(Baris).Cells(4).Value
            Dim durasi As TimeSpan = New TimeSpan
            durasi = now - start
            DataGridView1.Rows(Baris).Cells(5).Value = durasi
        Next
        rd.Close()
        conn.Close()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        For Baris As Integer = DataGridView2.Rows.Count - 1 To 0 Step -1
            'Label15.Text = rd("start_time")
            Dim now As DateTime = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            Dim start As DateTime = DataGridView2.Rows(Baris).Cells(5).Value
            Dim durasi As TimeSpan = New TimeSpan

            durasi = now - start
            DataGridView2.Rows(Baris).Cells(6).Value = durasi
            'Label16.Text = durasi.ToString

            If DataGridView2.Rows(Baris).Cells(6).Value >= "02:00:00" Then
                conn.Open()
                Dim sql As String
                sql = "UPDATE prodsys2_solder_paste_mgt_checkright_tbl SET run_time = @run_time, end_time = @end_time, status = @st WHERE txn_id = @txnid"
                cmd.Connection = conn
                cmd.CommandText = sql
                cmd.Parameters.AddWithValue("@txnid", DataGridView2.Rows(Baris).Cells(0).Value)
                cmd.Parameters.AddWithValue("@run_time", DataGridView2.Rows(Baris).Cells(6).Value)
                cmd.Parameters.AddWithValue("@end_time", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                cmd.Parameters.AddWithValue("@st", "EMPTY")
                Dim r As Integer
                r = cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()

                cmd = New MySql.Data.MySqlClient.MySqlCommand
                cmd.Connection = conn
                str = "INSERT INTO `prodsys2_solder_paste_mgt_mixer_tbl` (`txn_id`, `topbott` , `start_time`, `status`) VALUES (@id,@topbott,@start_time,@stat)"
                cmd.Parameters.Add("@id", MySqlDbType.VarChar).Value = DataGridView2.Rows(Baris).Cells(0).Value
                cmd.Parameters.Add("@topbott", MySqlDbType.VarChar).Value = DataGridView2.Rows(Baris).Cells(3).Value
                cmd.Parameters.Add("@start_time", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                cmd.Parameters.Add("@stat", MySqlDbType.VarChar).Value = "AVAILABLE"
                cmd.CommandText = str
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try

                Call LoadDataCoolingCase()
                Call LoadDataCheckRight()
                Call LoadDataMixer()
                Call LoadDataLine()
                Call LoadDataReadyMixer()
                Call LoadDataInLine()

                Call ShowCountCoolingCase()
                Call ShowCountCheckRight()
                Call ShowCountMixer()
                Call ShowCountLine()
                Call ShowCountReadyMixing()
                Call ShowCountInline()
                conn.Close()
            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormCheckRight.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call LoadDataCoolingCase()
        Call LoadDataCheckRight()
        Call LoadDataMixer()
        Call LoadDataLine()
        Call LoadDataReadyMixer()
        Call LoadDataInLine()
        Call LoadDataLogPemakaian()
        Call LoadDataPemakaianCheckright()
        Call LoadDataPemakaianMixer()
        Call LoadDataPemakaianLine()

        Call ShowCountCoolingCase()
        Call ShowCountCheckRight()
        Call ShowCountMixer()
        Call ShowCountLine()
        Call ShowCountReadyMixing()
        Call ShowCountInline()

        ComboBox1.SelectedItem = Nothing
        TextBox2.Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormMixer.ShowDialog()
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        For Baris As Integer = DataGridView3.Rows.Count - 1 To 0 Step -1
            'Label15.Text = rd("start_time")
            Dim now As DateTime = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            Dim start As DateTime = DataGridView3.Rows(Baris).Cells(4).Value
            Dim durasi As TimeSpan = New TimeSpan

            durasi = now - start
            DataGridView3.Rows(Baris).Cells(5).Value = durasi
            'Label16.Text = durasi.ToString

            If DataGridView3.Rows(Baris).Cells(5).Value > "00:00:45" Then
                conn.Open()
                Dim sql As String
                rd.Close()
                sql = "UPDATE prodsys2_solder_paste_mgt_mixer_tbl SET run_time = @run_time, end_time = @end_time, status = @status WHERE txn_id = @txn_id"
                cmd.Connection = conn
                cmd.CommandText = sql
                cmd.Parameters.AddWithValue("@txn_id", DataGridView3.Rows(Baris).Cells(0).Value)
                cmd.Parameters.AddWithValue("@run_time", DataGridView3.Rows(Baris).Cells(5).Value)
                cmd.Parameters.AddWithValue("@end_time", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                cmd.Parameters.AddWithValue("@status", "ALREADY MIXED")
                Dim r As Integer
                r = cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()

                Call LoadDataCoolingCase()
                Call LoadDataCheckRight()
                Call LoadDataMixer()
                Call LoadDataLine()
                Call LoadDataReadyMixer()
                Call LoadDataInLine()

                Call ShowCountCoolingCase()
                Call ShowCountCheckRight()
                Call ShowCountMixer()
                Call ShowCountLine()
                Call ShowCountReadyMixing()
                Call ShowCountInline()
                conn.Close()
            End If
        Next
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        For Baris As Integer = DataGridView4.Rows.Count - 1 To 0 Step -1
            Dim now As DateTime = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
            Dim start As DateTime = DataGridView4.Rows(Baris).Cells(5).Value
            Dim durasi As TimeSpan = New TimeSpan
            durasi = now - start
            DataGridView4.Rows(Baris).Cells(6).Value = durasi
        Next
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FormTrashSolderPaste.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        FormLine.ShowDialog()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If TextBox2.Text = "COOLING CASE IN" Then
                FormCoolingCase.ShowDialog()
                TextBox2.Clear()
            ElseIf TextBox2.Text = "CHECKRIGHT" Then
                FormCheckRight.ShowDialog()
                TextBox2.Clear()
            ElseIf TextBox2.Text = "MIXER" Then
                FormMixer.ShowDialog()
                TextBox2.Clear()
            ElseIf TextBox2.Text = "LINEIN" Then
                FormLine.ShowDialog()
                TextBox2.Clear()
            ElseIf TextBox2.Text = "TRASH" Then
                FormTrashSolderPaste.ShowDialog()
                TextBox2.Clear()
            Else
                MsgBox("KODE SALAH!", vbInformation)
                TextBox2.Clear()
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        FormSolderPasteReturn.ShowDialog()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If ComboBox1.Text = "COOLING CASE" Then
                Try
                    str = "SELECT txn_id, topbott, start_time, end_time, run_time FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE status = 'EMPTY' AND txn_id ='" & TextBox1.Text & "' ORDER BY end_time DESC"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView7.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call KolomLogPemakaian()
                TextBox1.Clear()
            ElseIf ComboBox1.Text = "CHECK RIGHT" Then
                Try
                    str = "SELECT txn_id, topbott, start_time, end_time, run_time, pic FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'EMPTY' AND txn_id ='" & TextBox1.Text & "' ORDER BY end_time DESC"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView8.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call KolomLogPemakaianCheckright()
                TextBox1.Clear()
            ElseIf ComboBox1.Text = "MIXER" Then
                Try
                    str = "SELECT txn_id, topbott, start_time, end_time, run_time FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE status = 'EMPTY' AND txn_id ='" & TextBox1.Text & "' ORDER BY end_time DESC"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView9.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call KolomLogPemakaianMixer()
                TextBox1.Clear()
            ElseIf ComboBox1.Text = "LINE" Then
                Try
                    str = "SELECT txn_id, topbott, line, start_time, end_time, run_time, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'EMPTY' AND txn_id ='" & TextBox1.Text & "' ORDER BY end_time DESC"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView10.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call KolomLogPemakaianLine()
                TextBox1.Clear()
            Else
                MsgBox("DATA TIDAK DITEMUKAN", vbInformation)
                TextBox1.Clear()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Select()
    End Sub
End Class