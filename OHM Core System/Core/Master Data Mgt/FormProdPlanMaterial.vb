Imports System.Security.Cryptography
Public Class FormProdPlanMaterial
    Private tripleDes As New TripleDESCryptoServiceProvider
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
            .ColumnHeadersHeight = 35
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.DarkCyan
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.DimGray
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
            .ColumnHeadersHeight = 35
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.DarkCyan
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.DimGray
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.SelectedCells(1).Value.ToString = "A2C" Then
            FormPreparelistMaterial.CheckBox1.Checked = True
            FormPreparelistMaterial.CheckBox2.Checked = False
        ElseIf DataGridView1.SelectedCells(1).Value.ToString = "A2F" Then
            FormPreparelistMaterial.CheckBox1.Checked = False
            FormPreparelistMaterial.CheckBox2.Checked = True
        End If
        FormPreparelistMaterial.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' wo supply
        FormPreparelistMaterial.TextBox3.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' pn
        FormPreparelistMaterial.TextBox4.Text = Me.DataGridView1.SelectedCells(7).Value.ToString ' qty
        FormPreparelistMaterial.RadioButton2.Checked = False
        FormPreparelistMaterial.RadioButton1.Checked = True
        FormPreparelistMaterial.TextBox2.Text = ""
        FormPreparelistMaterial.GroupBox1.Enabled = True
        FormPreparelistMaterial.GroupBox2.Enabled = True
        FormPreparelistMaterial.GroupBox3.Enabled = True
        Me.Close()
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        FormPreparelistMaterial.TextBox3.Text = Me.DataGridView2.SelectedCells(0).Value.ToString ' pn
        FormPreparelistMaterial.TextBox2.Text = Me.DataGridView2.SelectedCells(3).Value.ToString ' wo demand
        FormPreparelistMaterial.TextBox4.Text = Me.DataGridView2.SelectedCells(2).Value.ToString ' qty


        FormPreparelistMaterial.RadioButton1.Checked = False
        FormPreparelistMaterial.RadioButton2.Checked = True
        FormPreparelistMaterial.TextBox1.Text = ""
        FormPreparelistMaterial.GroupBox1.Enabled = True
        FormPreparelistMaterial.GroupBox2.Enabled = True
        FormPreparelistMaterial.GroupBox3.Enabled = True
        Me.Close()
    End Sub

    Private Sub FormProdPlanMaterial_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call customdgv2()

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Me.TabControl1.SelectedIndex = 0 Then ' SMT PLAN
            If ComboBox1.Text = "Supply WO" Then
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where supply_wo like '" & TextBox1.Text & "' ", conn)
                ds = New DataSet
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
            ElseIf ComboBox1.Text = "Partnumber" Then
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where pn like '" & TextBox1.Text & "' ", conn)
                ds = New DataSet
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
            End If
        ElseIf Me.TabControl1.SelectedIndex = 1 Then ' ASP
            If ComboBox1.Text = "Demand WO" Then
                DataGridView2.DataSource = Nothing
                DataGridView2.Refresh()
                da = New MySql.Data.MySqlClient.MySqlDataAdapter("select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where wo like '" & TextBox1.Text & "' ", conn)
                ds = New DataSet
                da.Fill(ds)
                DataGridView2.DataSource = ds.Tables(0)
            ElseIf ComboBox1.Text = "Partnumber" Then
                DataGridView2.DataSource = Nothing
                DataGridView2.Refresh()
                da = New MySql.Data.MySqlClient.MySqlDataAdapter("select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where ebt like '" & TextBox1.Text & "' ", conn)
                ds = New DataSet
                da.Fill(ds)
                DataGridView2.DataSource = ds.Tables(0)
            End If
        End If
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

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Me.TabControl1.SelectedIndex = 0 Then ' SMT PLAN
                If ComboBox1.Text = "Supply WO" Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where supply_wo like '" & TextBox1.Text & "' ", conn)
                    ds = New DataSet
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                ElseIf ComboBox1.Text = "Partnumber" Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where pn like '" & TextBox1.Text & "' ", conn)
                    ds = New DataSet
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                End If
            ElseIf Me.TabControl1.SelectedIndex = 1 Then ' ASP
                If ComboBox1.Text = "Demand WO" Then
                    DataGridView2.DataSource = Nothing
                    DataGridView2.Refresh()
                    da = New MySql.Data.MySqlClient.MySqlDataAdapter("select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where wo like '" & TextBox1.Text & "' ", conn)
                    ds = New DataSet
                    da.Fill(ds)
                    DataGridView2.DataSource = ds.Tables(0)
                ElseIf ComboBox1.Text = "Partnumber" Then
                    DataGridView2.DataSource = Nothing
                    DataGridView2.Refresh()
                    da = New MySql.Data.MySqlClient.MySqlDataAdapter("select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where ebt like '" & TextBox1.Text & "' ", conn)
                    ds = New DataSet
                    da.Fill(ds)
                    DataGridView2.DataSource = ds.Tables(0)
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.Text IsNot Nothing Then
            TextBox1.Focus()
        End If
    End Sub
End Class