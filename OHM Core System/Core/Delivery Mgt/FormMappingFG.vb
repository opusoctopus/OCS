Imports MySql.Data.MySqlClient
Public Class FormMappingFG

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
        'FormMappingFG.DataGridView1.Columns(0).Width = 100 'txnid
        DataGridView1.Columns(0).Width = 100 'locator
        DataGridView1.Columns(1).Width = 100 'wo
        DataGridView1.Columns(2).Width = 120 'pn
        DataGridView1.Columns(3).Width = 160 'model
        DataGridView1.Columns(4).Width = 100 'demand_date
        DataGridView1.Columns(5).Width = 60 'supply_qty
        DataGridView1.Columns(6).Width = 60 'inspect_qty
        DataGridView1.Columns(7).Width = 70 'remained_qty
        DataGridView1.Columns(8).Width = 180 'date_in
        DataGridView1.Columns(9).Width = 182 'date_out
        'DataGridView1.Columns(11).Width = 180 'percent
    End Sub

    Sub BuatKolom2()
        'FormMappingFG.DataGridView1.Columns(0).Width = 100 'txnid
        DataGridView2.Columns(0).Width = 100 'locator
        DataGridView2.Columns(1).Width = 80 'wo
        DataGridView2.Columns(2).Width = 100 'pn
        DataGridView2.Columns(3).Width = 100 'pst
        DataGridView2.Columns(4).Width = 140 'qty
    End Sub

    Sub LoadData()
        Try
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, model as Model, demand_date as PST, supply_qty as Qty, inspect_qty as OQA, remained_qty as Remain, date_in as DateIn, date_out as DateOut FROM prodsys2_whfg_in_status_tbl WHERE locator IS NULL ORDER BY locator ASC"
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

    Sub LoadData2()
        Try
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, demand_date as PST, supply_qty as Qty FROM prodsys2_whfg_in_status_tbl WHERE date_out IS NULL ORDER BY locator ASC"
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

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
        Label2.Text = DataGridView2.RowCount
    End Sub

    Sub KolomDashboard()
        FormLocatorFG.DataGridView1.Columns(0).Width = 160 'locator
        FormLocatorFG.DataGridView1.Columns(1).Width = 170 'wo
        FormLocatorFG.DataGridView1.Columns(2).Width = 210 'pn
        FormLocatorFG.DataGridView1.Columns(3).Width = 350 'model
        FormLocatorFG.DataGridView1.Columns(4).Width = 170 'demand_date
        FormLocatorFG.DataGridView1.Columns(5).Width = 100 'supply_qty
        FormLocatorFG.DataGridView1.Columns(6).Width = 100 'inspect_qty
        FormLocatorFG.DataGridView1.Columns(7).Width = 130 'remained_qty
        FormLocatorFG.DataGridView1.Columns(8).Width = 265 'date_in
        FormLocatorFG.DataGridView1.Columns(9).Width = 265 'date_out
    End Sub

    Sub LoadDataDashboard()
        Try
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, model as Model, demand_date as PST, supply_qty as Qty, inspect_qty as OQA, remained_qty as Remain, date_in as DateIn, date_out as DateOut FROM prodsys2_whfg_in_status_tbl WHERE locator IS NOT NULL AND date_out IS NULL ORDER BY locator ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormLocatorFG.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomDashboard()
    End Sub

    Private Sub FormMappingFG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call LoadData2()
        Call showcount()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormFGIn.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        FormFGOut.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormMappingLocator.TextBox2.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' WO
        FormMappingLocator.TextBox3.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' PN
        FormMappingLocator.TextBox4.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' MODEL
        FormMappingLocator.ComboBox1.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' Locator

        FormMappingLocator.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim screen2 As Screen
        screen2 = Screen.AllScreens(0)
        FormLocatorFG.StartPosition = FormStartPosition.Manual
        FormLocatorFG.Location = screen2.Bounds.Location + New Point(100, 100)
        FormLocatorFG.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormLocatorFG.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call LoadDataDashboard()
        Call KolomDashboard()
    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        FormMappingLocatorUpdate.TextBox2.Text = Me.DataGridView2.SelectedCells(1).Value.ToString ' WO
        FormMappingLocatorUpdate.TextBox3.Text = Me.DataGridView2.SelectedCells(2).Value.ToString ' PN
        FormMappingLocatorUpdate.ComboBox1.Text = Me.DataGridView2.SelectedCells(0).Value.ToString ' Locator

        FormMappingLocatorUpdate.ShowDialog()
    End Sub
End Class