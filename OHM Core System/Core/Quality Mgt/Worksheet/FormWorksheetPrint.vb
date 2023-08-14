Imports Microsoft.Reporting.WinForms
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient

Public Class FormWorksheetPrint
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
        ComboBox1.SelectedItem = Nothing
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
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
    End Sub

    Sub BuatKolomSub()
        DataGridView2.Columns(0).Width = 70 'Barcode
        DataGridView2.Columns(1).Width = 180 'Barcode
        DataGridView2.Columns(2).Width = 130 'Barcode
        DataGridView2.Columns(3).Width = 220 'Barcode
    End Sub

    Sub BuatKolomSubSelected()
        DataGridView3.Columns(0).Width = 150 'Barcode
        DataGridView3.Columns(1).Width = 410 'Barcode
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormProdPlan2.ShowDialog()
    End Sub

    Private Sub FormWorksheetPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        GroupBox1.Enabled = False
        RadioButton3.Enabled = False

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

    Private Sub deleteData(sql As String)
        Dim result As Integer
        Try
            Call connection()
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                result = .ExecuteNonQuery()
            End With
            If result > 0 Then
                'MsgBox("Data has been deleted in the database")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, child, description, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox1.Text & "' AND parent_desc='BPR Insert PCB Assembly' AND supply_type='Assembly Pull'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT parent_desc, child, description FROM prodsys2_master_bom_tbl WHERE (child LIKE 'EAG%' OR child LIKE 'EBL%' OR child LIKE '6%' AND description NOT LIKE 'DIN') AND id_bom='" & TextBox1.Text & "' AND level NOT IN (1,2,5,6,7) AND `description` NOT LIKE '%DIN%' AND parent_desc = 'Substitutes'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)
        Call BuatKolomSub()

        Dim sql As String
        Call connection()
        str = "SELECT child_sub, description_sub FROM prodsys2_master_bom_sub_tbl WHERE id_bom_sub='" & TextBox1.Text & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                sql = "DELETE FROM prodsys2_master_bom_sub_tbl WHERE id_bom_sub='" & TextBox1.Text & "'"
                deleteData(sql)

                rd.Close()
                Call connection()
                Dim cmd As MySqlCommand
                For i As Integer = 0 To DataGridView2.Rows.Count - 1 Step +1
                    cmd = New MySqlCommand("INSERT INTO `prodsys2_master_bom_sub_tbl`(`id_bom_sub`, `parent_desc_sub`, `child_sub`, `description_sub`, `updated`) VALUES (@id_bom_sub, @parent_desc_sub, @child_sub, @description_sub, @updated)", conn)
                    cmd.Parameters.Add("@id_bom_sub", MySqlDbType.VarChar).Value = TextBox1.Text
                    cmd.Parameters.Add("@parent_desc_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(1).Value.ToString()
                    cmd.Parameters.Add("@child_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(2).Value.ToString()
                    cmd.Parameters.Add("@description_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(3).Value.ToString()
                    cmd.Parameters.Add("@updated", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.ExecuteNonQuery()
                Next
                conn.Close()
                'MessageBox.Show("All Data Updated")
            Else
                rd.Close()
                Call connection()
                Dim cmd As MySqlCommand
                For i As Integer = 0 To DataGridView2.Rows.Count - 1 Step +1
                    cmd = New MySqlCommand("INSERT INTO `prodsys2_master_bom_sub_tbl`(`id_bom_sub`, `parent_desc_sub`, `child_sub`, `description_sub`, `updated`) VALUES (@id_bom_sub, @parent_desc_sub, @child_sub, @description_sub, @updated)", conn)
                    cmd.Parameters.Add("@id_bom_sub", MySqlDbType.VarChar).Value = TextBox1.Text
                    cmd.Parameters.Add("@parent_desc_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(1).Value.ToString()
                    cmd.Parameters.Add("@child_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(2).Value.ToString()
                    cmd.Parameters.Add("@description_sub", MySqlDbType.VarChar).Value = DataGridView2.Rows(i).Cells(3).Value.ToString()
                    cmd.Parameters.Add("@updated", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.ExecuteNonQuery()
                Next
                conn.Close()
                'MessageBox.Show("All Data Inserted")
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try

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

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_worksheet_tbl`(`id_worksheet`, `process`, `pn`, `supply_wo` , `demand_wo`, `lot_qty`, `line`, `printed`) VALUES (@id_worksheet,@process,@pn,@supply_wo,@demand_wo,@lot_qty,@line,@printed)"
        cmd.Parameters.Add("@id_worksheet", MySqlDbType.VarChar).Value = TextBox5.Text
        cmd.Parameters.Add("@process", MySqlDbType.VarChar).Value = TextBox6.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@demand_wo", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@lot_qty", MySqlDbType.VarChar).Value = TextBox4.Text
        cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@printed", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()

            FormWorksheetPreview.ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
            FormWorksheetPreview.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.FullPage
            FormWorksheetPreview.ShowDialog()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Call bersih()
        GroupBox1.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call bersih()
        GroupBox1.Enabled = False
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