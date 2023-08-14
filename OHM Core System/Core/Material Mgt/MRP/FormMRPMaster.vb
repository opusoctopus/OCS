Imports MySql.Data.MySqlClient

Public Class FormMRPMaster

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
    End Sub

    Private Sub FormMRPMaster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Call customdgv2()
            str = "SELECT * FROM prodsys2_master_mrp_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try

        Dim command As New MySqlCommand("select distinct supply_wo from prodsys2_master_mrp_tbl", conn)
        Dim adapter As New MySqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox1.DataSource = table
        ComboBox1.DisplayMember = "supply_wo"
        'ComboBox1.ValueMember = "Id"

        Dim command2 As New MySqlCommand("select distinct level from prodsys2_master_mrp_tbl", conn)
        Dim adapter2 As New MySqlDataAdapter(command2)
        Dim table2 As New DataTable()
        adapter2.Fill(table2)
        ComboBox2.DataSource = table2
        ComboBox2.DisplayMember = "level"
        'ComboBox1.ValueMember = "Id"

        Label2.Text = DataGridView1.RowCount
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            DataGridView1.DataSource.clear
            DataGridView1.Refresh()
            str = "SELECT * FROM prodsys2_master_mrp_tbl WHERE supply_wo = '" & ComboBox1.Text & "' AND level = '" & ComboBox2.Text & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
            Label2.Text = DataGridView1.RowCount
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Call customdgv2()
            str = "SELECT * FROM prodsys2_master_mrp_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try

        Dim command As New MySqlCommand("select distinct supply_wo from prodsys2_master_mrp_tbl", conn)
        Dim adapter As New MySqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox1.DataSource = table
        ComboBox1.DisplayMember = "supply_wo"
        'ComboBox1.ValueMember = "Id"

        Dim command2 As New MySqlCommand("select distinct level from prodsys2_master_mrp_tbl", conn)
        Dim adapter2 As New MySqlDataAdapter(command2)
        Dim table2 As New DataTable()
        adapter2.Fill(table2)
        ComboBox2.DataSource = table2
        ComboBox2.DisplayMember = "level"
        'ComboBox1.ValueMember = "Id"

        Label2.Text = DataGridView1.RowCount
    End Sub
End Class