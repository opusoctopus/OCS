Imports Microsoft.Reporting.WinForms
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient

Public Class FormBOMKit
    Public Property Truerow As Boolean
    Private tripleDes As New TripleDESCryptoServiceProvider

    Sub bersih()
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        ComboBox1.Items.Clear()
        ComboBox1.SelectedItem = Nothing
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
    End Sub

    Private Sub FormBOMKit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        GroupBox1.Enabled = False
        RadioButton3.Enabled = False
        Button2.Enabled = False

        Dim checkboxcolumn As New DataGridViewCheckBoxColumn
        checkboxcolumn.Width = 70
        checkboxcolumn.Name = "checkboxcolumn"
        checkboxcolumn.HeaderText = "Exclude"
        DataGridView1.Columns.Insert(0, checkboxcolumn)

        Dim checkboxcolumn2 As New DataGridViewCheckBoxColumn
        checkboxcolumn2.Width = 70
        checkboxcolumn2.Name = "checkboxcolumn2"
        checkboxcolumn2.HeaderText = "Exclude"
        DataGridView2.Columns.Insert(0, checkboxcolumn2)
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(112, 122, 131)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With

        With DataGridView2 'Ganti dengan nama datagridview kalian
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(112, 122, 131)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With

        With DataGridView3 'Ganti dengan nama datagridview kalian
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(112, 122, 131)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormProdPlan2.ShowDialog()
    End Sub

    Sub BuatKolomSub()
        DataGridView2.Columns(0).Width = 35 'Barcode
        DataGridView2.Columns(1).Width = 150 'Barcode
        DataGridView2.Columns(2).Width = 350 'Barcode
    End Sub

    Sub BuatKolomSubSelected()
        DataGridView3.Columns(0).Width = 150 'Barcode
        DataGridView3.Columns(1).Width = 410 'Barcode
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged 'BPR
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT * FROM prodsys2_master_bom_tbl WHERE (child LIKE 'EAG%' OR child LIKE 'EBL%' OR child LIKE '6%' AND description NOT LIKE 'DIN') AND id_bom='" & TextBox1.Text & "' AND level NOT IN (1,2,5,6,7) AND `description` NOT LIKE '%DIN%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT child as childsub, description as descriptionsub FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND parent_desc = ('Substitutes')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)
        Call BuatKolomSub()

        TextBox5.Text = EncryptData(TextBox1.Text)
        TextBox6.Text = ""
        TextBox6.Text = RadioButton1.Text
        ComboBox1.Items.Clear()

        With ComboBox1
            .Items.Add("BPR Line 1")
            .Items.Add("BPR Line 2")
            .Items.Add("BPR Line 3")
            .Items.Add("BPR Line 4")
            .Items.Add("BPR Line 5")
            .Items.Add("BPR Line 6")
        End With
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged 'DMS
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, child, description, specification, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND supply_type = ('Assembly Pull') AND level in (1,2) AND qty > 0 AND description NOT IN ('Label', 'Tape,Double Faced', 'Ribbon', 'PCB Assembly,Sub')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, parent_desc, child, specification, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND parent_desc = ('Substitutes')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)

        TextBox5.Text = EncryptData(TextBox1.Text)
        TextBox6.Text = ""
        TextBox6.Text = RadioButton2.Text
        ComboBox1.Items.Clear()
        With ComboBox1
            .Items.Add("DMS Offline")
            .Items.Add("DMS Inline 1")
            .Items.Add("DMS Inline 2")
            .Items.Add("DMS Inline 3")
            .Items.Add("DMS Inline 4")
            .Items.Add("DMS Inline 5")
            .Items.Add("DMS Inline 6")
        End With
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        'MsgBox("STATUS: WAIT POR EM . . . .!!", vbInformation)
        'DataGridView1.DataSource = Nothing
        'DataGridView1.Refresh()
        'DataGridView2.DataSource = Nothing
        'DataGridView2.Refresh()
        'da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, child, description, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND supply_type = ('Assembly Pull') AND level in (3,4) AND designators > ''", conn)
        'ds = New DataSet
        'da.Fill(ds)
        'DataGridView1.DataSource = ds.Tables(0)

        'rd.Close()
        'da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, parent_desc, child, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND parent_desc = ('Substitutes')", conn)
        'ds = New DataSet
        'da.Fill(ds)
        'DataGridView2.DataSource = ds.Tables(0)

        'TextBox5.Text = EncryptData(TextBox1.Text)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        CheckBox2.Checked = False
        TextBox7.Text = ""
        TextBox7.Text = CheckBox1.Text
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckBox1.Checked = False
        TextBox7.Text = ""
        TextBox7.Text = CheckBox2.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_worksheet_tbl`(`id_worksheet`, `process`, `pn`, `supply_wo` , `demand_wo`, `lot_qty`, `line`, `shift`, `printed`) VALUES (@id_worksheet,@process,@pn,@supply_wo,@demand_wo,@lot_qty,@line,@shift,@printed)"
        cmd.Parameters.Add("@id_worksheet", MySqlDbType.VarChar).Value = TextBox5.Text
        cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = TextBox6.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@demand_wo", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@lot_qty", MySqlDbType.VarChar).Value = TextBox4.Text
        cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@shift", MySqlDbType.VarChar).Value = TextBox7.Text
        cmd.Parameters.Add("@printed", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If RadioButton1.Checked = True Then
            Dim str As String = "SELECT * FROM prodsys2_master_bom_tbl WHERE (child LIKE 'EAG%' OR child LIKE 'EBL%' OR child LIKE '6%' AND description NOT LIKE 'DIN') AND id_bom='" & TextBox1.Text & "' AND level NOT IN (1,2,5,6,7) AND `description` NOT LIKE '%DIN%'"
            Try
                FormWorksheet.Worksheet.Clear()
                FormWorksheet.Worksheet.EnforceConstraints = False
                connection()
                da = New MySqlDataAdapter(str, conn)
                da.Fill(FormWorksheet.Worksheet.DataTable1)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            FormWorksheet.ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
            FormWorksheet.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent
            FormWorksheet.ReportViewer1.ZoomPercent = 25
            FormWorksheet.ReportViewer1.RefreshReport()
            FormWorksheet.Show()

        ElseIf RadioButton2.Checked = True Then
            Dim selectedRows2 As List(Of DataGridViewRow) = (From row In DataGridView1.Rows.Cast(Of DataGridViewRow)() Where Convert.ToBoolean(row.Cells("checkboxcolumn").Value) = Truerow).ToList()
            Dim worksheetdms As New FormWorksheetDMS(selectedRows2)
            worksheetdms.ShowDialog()
        Else
            MsgBox("NO WORKSHEET SELECTED", MsgBoxStyle.Information)
        End If
        Call bersih()
        GroupBox1.Enabled = False
    End Sub

    Public Function EncryptData(ByVal plaintext As String) As String
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
        Dim ms As New System.IO.MemoryStream
        Dim encStream As New CryptoStream(ms, tripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        Return Convert.ToBase64String(ms.ToArray)
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call bersih()
        GroupBox1.Enabled = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim table2 As New DataTable
        With table2
            .Columns.Add("childsub")
            .Columns.Add("descriptionsub")
        End With

        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim isselect As Boolean = Convert.ToBoolean(row.Cells("checkboxcolumn2").Value)
            If isselect Then
                table2.Rows.Add(row.Cells(1).Value, row.Cells(2).Value)
            End If
        Next
        DataGridView3.DataSource = table2
        Call BuatKolomSubSelected()
    End Sub
End Class