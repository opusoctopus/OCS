Imports MySql.Data.MySqlClient

Public Class FormPreparelistMaterial
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        Label9.Text = ""
        Label11.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        DataGridView3.DataSource = Nothing
        DataGridView3.Refresh()
        Button3.Enabled = False
    End Sub

    Sub cancel()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        Label9.Text = ""
        Label11.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        Button3.Enabled = False
    End Sub

    Sub add()
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Prepare", "Prepare")
        DataGridView1.Columns.Add("Supply WO", "Supply WO")
        DataGridView1.Columns.Add("Demand WO", "Demand WO")
        DataGridView1.Columns.Add("Partnumber", "Partnumber")
        DataGridView1.Columns.Add("Lot. QTY", "Lot. QTY")
        DataGridView1.Columns.Add("Prepare List", "Prepare List")
        DataGridView1.Columns.Add("Label Material", "Label Material")
        DataGridView1.Columns.Add("Result", "Result")
        DataGridView1.Columns.Add("Remark", "Remark")
        DataGridView1.Columns.Add("Created", "Created")
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 150
        DataGridView1.Columns(2).Width = 150
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).Width = 150
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(8).Width = 100
        DataGridView1.Columns(9).Width = 505
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

    Sub LoadBOM34()
        DataGridView2.DataSource = Nothing
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Refresh()

        DataGridView3.DataSource = Nothing
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        DataGridView3.Refresh()

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, child, description, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox3.Text & "' AND supply_type = ('Assembly Pull') AND level in (3,4) AND designators > '' AND parent_desc NOT IN ('Auto SMT PCB Assembly', 'Package Assembly')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, parent_desc, child, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox3.Text & "' AND parent_desc = ('Substitutes')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView3.DataSource = ds.Tables(0)
    End Sub

    Sub LoadBOM12()
        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, child, description, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox3.Text & "' AND supply_type = ('Assembly Pull') AND level in (1,2) AND qty > 0 AND description NOT IN ('Label', 'Tape,Double Faced', 'Ribbon', 'Primary Cell Battery,Lithium', 'PCB Assembly,Sub', 'Connector,DSUB', 'Tape,Polyester', 'Screw Assembly', 'Filter,AC Line', 'Screw,Taptite', 'Bag')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)

        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT id_bom, parent_desc, child, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & TextBox3.Text & "' AND parent_desc = ('Substitutes')", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView3.DataSource = ds.Tables(0)
    End Sub

    Sub LoadMRPBPR()
        'MAIN COMPONENT
        DataGridView2.DataSource = Nothing
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Refresh()

        DataGridView3.DataSource = Nothing
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        DataGridView3.Refresh()

        If CheckBox1.Checked = True Then 'A2C
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='BPR' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox1.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf CheckBox2.Checked = True Then 'A2F
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='BPR' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox2.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        Else
            MsgBox("ORG not checked!", MsgBoxStyle.Information)
        End If
    End Sub

    Sub LoadMRPSMT()
        'MAIN COMPONENT
        DataGridView2.DataSource = Nothing
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Refresh()

        DataGridView3.DataSource = Nothing
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        DataGridView3.Refresh()

        If CheckBox1.Checked = True Then 'A2C
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='SMT' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox1.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf CheckBox2.Checked = True Then 'A2F
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='SMT' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox2.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        Else
            MsgBox("ORG not checked!", MsgBoxStyle.Information)
        End If
    End Sub

    Sub LoadMRPDMS()
        'MAIN COMPONENT
        DataGridView2.DataSource = Nothing
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Refresh()

        DataGridView3.DataSource = Nothing
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        DataGridView3.Refresh()

        If CheckBox1.Checked = True Then 'A2C
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='DMS' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox1.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        ElseIf CheckBox2.Checked = True Then 'A2F
            Try
                str = "SELECT supply_wo, child, description, qpa FROM prodsys2_master_mrp_tbl WHERE level='DMS' AND supply_wo = '" & TextBox1.Text & "' AND qpa >= 1 AND org='" & CheckBox2.Text & "'"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView2.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
        Else
            MsgBox("ORG not checked!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub FormPreparelistMaterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call bersih()
        Call BuatKolom()
        Call customdgv2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call add()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            TextBox1.Enabled = True
            TextBox1.Select()
            'TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label9.Text = ""
            Label11.Text = ""
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            TextBox2.Enabled = True
            TextBox2.Select()
            'TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label9.Text = ""
            Label11.Text = ""
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            TextBox1.Enabled = True
            TextBox1.Select()
            'TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label9.Text = ""
            Label11.Text = ""
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If CheckBox1.Checked = False And CheckBox2.Checked = False Then
                MsgBox("ORG [" & CheckBox1.Text & "] & [" & CheckBox2.Text & "] not checked!", vbCritical)
            ElseIf CheckBox1.Checked = True Then
                Call connection()
                str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox1.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If Not rd.HasRows Then
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "Production plan hasn't been updated yet"
                        Label11.Font = New Font("Arial Narrow", 14)
                        Label11.ForeColor = Color.Crimson
                        TextBox1.Text = ""
                    Else
                        Label9.Text = ""
                        Label11.Text = ""
                        TextBox1.Text = rd("supply_wo")
                        TextBox2.Text = rd("demand_wo")
                        TextBox3.Text = rd("pn")
                        TextBox4.Text = rd("lot_qty")
                        TextBox5.Focus()

                        Call LoadBOM34()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    rd.Close()
                End Try
            ElseIf CheckBox2.Checked = True Then
                Call connection()
                str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If Not rd.HasRows Then
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "Production plan hasn't been updated yet"
                        Label11.Font = New Font("Arial Narrow", 14)
                        Label11.ForeColor = Color.Crimson
                        TextBox1.Text = ""
                    Else
                        Label9.Text = ""
                        Label11.Text = ""
                        TextBox1.Text = rd("supply_wo")
                        TextBox2.Text = rd("demand_wo")
                        TextBox3.Text = rd("pn")
                        TextBox4.Text = rd("lot_qty")
                        TextBox5.Focus()

                        Call LoadBOM34()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    rd.Close()
                End Try
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call cancel()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        'TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox1.Select()
        Label9.Text = ""
        Label11.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        CheckBox2.Checked = False
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox1.Select()
        Label9.Text = ""
        Label11.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        CheckBox1.Checked = False
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Call connection()
            str = "SELECT ebt, supply_wo, wo, wo_qty FROM prodsys_prod_plan_tbl WHERE wo='" & TextBox2.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If Not rd.HasRows Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "Production plan hasn't been updated yet"
                    Label11.Font = New Font("Arial Narrow", 14)
                    Label11.ForeColor = Color.Crimson
                    TextBox2.Text = ""
                Else
                    Label9.Text = ""
                    Label11.Text = ""
                    TextBox1.Text = rd("supply_wo")
                    TextBox2.Text = rd("wo")
                    TextBox3.Text = rd("ebt")
                    TextBox4.Text = rd("wo_qty")
                    TextBox5.Focus()

                    Call LoadBOM12()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                rd.Close()
            End Try
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Text = TextBox5.Text.Substring(3)
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            If RadioButton1.Checked = True Then ' =============== BPR
                If TextBox5.Text = "" And TextBox6.Text = "" Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "NO BARCODE IS SCANNED"
                    Label11.Font = New Font("Arial Narrow", 14)
                    Label11.ForeColor = Color.Crimson
                ElseIf TextBox5.Text <> TextBox6.Text Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "WRONG COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                    Label11.Font = New Font("Arial Narrow", 12)
                    Label11.ForeColor = Color.Crimson
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox5.Focus()
                    'ElseIf TextBox5.Text = "EAG37842019" Or TextBox6.Text = "EAG37842019" Then
                    'Label9.Text = "NG"
                    'Label9.Font = New Font("Arial", 72)
                    'Label9.ForeColor = Color.Red
                    'Label11.Text = "MAKER SPG, COMPONENT [" & TextBox6.Text & "] STATUS HOLD!"
                    'Label11.Font = New Font("Arial Narrow", 12)
                    'Label11.ForeColor = Color.Crimson
                    'TextBox5.Text = ""
                    'TextBox6.Text = ""
                    'TextBox5.Focus()
                Else
                    Dim IsFound As Boolean = False
                    For Each row As DataGridViewRow In DataGridView2.Rows
                        For i As Integer = 0 To DataGridView2.Columns.Count - 1
                            If row.Cells(i).Value.ToString = TextBox6.Text Then
                                IsFound = True
                                row.Cells(i).Style.BackColor = Color.Lime
                                row.Cells(i).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next
                    For Each row As DataGridViewRow In DataGridView3.Rows
                        For j As Integer = 0 To DataGridView3.Columns.Count - 1
                            If row.Cells(j).Value.ToString = TextBox6.Text Then
                                IsFound = True
                                row.Cells(j).Style.BackColor = Color.Lime
                                row.Cells(j).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next

                    If (IsFound) Then
                        Label11.Text = ""
                        Label9.Text = "OK"
                        Label9.ForeColor = Color.Lime
                        Label9.Font = New Font("Arial", 72)

                        Dim duplicate As Boolean = False
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                If row.Cells(i).Value.ToString = TextBox6.Text Then
                                    duplicate = True
                                End If
                            Next
                        Next

                        If (duplicate) Then
                            Label9.Text = "NG"
                            Label9.Font = New Font("Arial", 72)
                            Label9.ForeColor = Color.Red
                            Label11.Text = "DUPLICATE ENTRY [" & TextBox6.Text & "]"
                            Label11.Font = New Font("Arial Narrow", 12)
                            Label11.ForeColor = Color.Crimson
                        Else
                            DataGridView1.Rows.Add(New String() {RadioButton1.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, Label9.Text, "OK PREPARE", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                        End If

                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                        If DataGridView1.RowCount <> DataGridView2.RowCount Then
                            Button3.Enabled = False
                        Else : DataGridView1.RowCount = DataGridView2.RowCount
                            Button3.Enabled = True
                            Exit Sub
                        End If
                    Else
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "MIXING COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                        Label11.Font = New Font("Arial Narrow", 12)
                        Label11.ForeColor = Color.Crimson
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                    End If
                End If
            ElseIf RadioButton2.Checked = True Then ' =========== DMS
                If TextBox5.Text = "" And TextBox6.Text = "" Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "NO BARCODE IS SCANNED"
                    Label11.Font = New Font("Arial Narrow", 14)
                    Label11.ForeColor = Color.Crimson
                ElseIf TextBox5.Text <> TextBox6.Text Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "WRONG COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                    Label11.Font = New Font("Arial Narrow", 12)
                    Label11.ForeColor = Color.Crimson
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox5.Focus()
                Else
                    Dim IsFound As Boolean = False
                    For Each row As DataGridViewRow In DataGridView2.Rows
                        For i As Integer = 0 To DataGridView2.Columns.Count - 1
                            If row.Cells(i).Value.ToString = TextBox5.Text Then
                                IsFound = True
                                row.Cells(i).Style.BackColor = Color.Lime
                                row.Cells(i).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next
                    For Each row As DataGridViewRow In DataGridView3.Rows
                        For j As Integer = 0 To DataGridView3.Columns.Count - 1
                            If row.Cells(j).Value.ToString = TextBox6.Text Then
                                IsFound = True
                                row.Cells(j).Style.BackColor = Color.Lime
                                row.Cells(j).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next
                    If (IsFound) Then
                        Label11.Text = ""
                        Label9.Text = "OK"
                        Label9.ForeColor = Color.Lime
                        Label9.Font = New Font("Arial", 72)
                        Dim duplicate As Boolean = False
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                If row.Cells(i).Value.ToString = TextBox6.Text Then
                                    duplicate = True
                                End If
                            Next
                        Next

                        If (duplicate) Then
                            Label9.Text = "NG"
                            Label9.Font = New Font("Arial", 72)
                            Label9.ForeColor = Color.Red
                            Label11.Text = "DUPLICATE ENTRY [" & TextBox6.Text & "]"
                            Label11.Font = New Font("Arial Narrow", 12)
                            Label11.ForeColor = Color.Crimson
                        Else
                            DataGridView1.Rows.Add(New String() {RadioButton2.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, Label9.Text, "OK", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                        End If
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                        If DataGridView1.RowCount <> DataGridView2.RowCount Then
                            Button3.Enabled = False
                        Else : DataGridView1.RowCount = DataGridView2.RowCount
                            Button3.Enabled = True
                            Exit Sub
                        End If
                    Else
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "MIXING COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                        Label11.Font = New Font("Arial Narrow", 14)
                        Label11.ForeColor = Color.Crimson
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                    End If
                End If
            ElseIf RadioButton3.Checked = True Then ' =========== SMT
                If TextBox5.Text = "" And TextBox6.Text = "" Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "NO BARCODE IS SCANNED"
                    Label11.Font = New Font("Arial Narrow", 14)
                    Label11.ForeColor = Color.Crimson
                ElseIf TextBox5.Text <> TextBox6.Text Then
                    Label9.Text = "NG"
                    Label9.Font = New Font("Arial", 72)
                    Label9.ForeColor = Color.Red
                    Label11.Text = "WRONG COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                    Label11.Font = New Font("Arial Narrow", 12)
                    Label11.ForeColor = Color.Crimson
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox5.Focus()
                Else
                    Dim IsFound As Boolean = False
                    For Each row As DataGridViewRow In DataGridView2.Rows
                        For i As Integer = 0 To DataGridView2.Columns.Count - 1
                            If row.Cells(i).Value.ToString = TextBox5.Text Then
                                IsFound = True
                                row.Cells(i).Style.BackColor = Color.Lime
                                row.Cells(i).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next
                    For Each row As DataGridViewRow In DataGridView3.Rows
                        For j As Integer = 0 To DataGridView3.Columns.Count - 1
                            If row.Cells(j).Value.ToString = TextBox6.Text Then
                                IsFound = True
                                row.Cells(j).Style.BackColor = Color.Lime
                                row.Cells(j).Style.ForeColor = Color.Black
                            End If
                        Next
                    Next
                    If (IsFound) Then
                        Label11.Text = ""
                        Label9.Text = "OK"
                        Label9.ForeColor = Color.Lime
                        Label9.Font = New Font("Arial", 72)
                        Dim duplicate As Boolean = False
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                If row.Cells(i).Value.ToString = TextBox6.Text Then
                                    duplicate = True
                                End If
                            Next
                        Next

                        If (duplicate) Then
                            Label9.Text = "NG"
                            Label9.Font = New Font("Arial", 72)
                            Label9.ForeColor = Color.Red
                            Label11.Text = "DUPLICATE ENTRY [" & TextBox6.Text & "]"
                            Label11.Font = New Font("Arial Narrow", 12)
                            Label11.ForeColor = Color.Crimson
                        Else
                            DataGridView1.Rows.Add(New String() {RadioButton3.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, Label9.Text, "OK", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")})
                        End If
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                        If DataGridView1.RowCount <> DataGridView2.RowCount Then
                            Button3.Enabled = False
                        Else : DataGridView1.RowCount = DataGridView2.RowCount
                            Button3.Enabled = True
                            Exit Sub
                        End If
                    Else
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "MIXING COMPONENT [" & TextBox5.Text & "] & [" & TextBox6.Text & "]"
                        Label11.Font = New Font("Arial Narrow", 14)
                        Label11.ForeColor = Color.Crimson
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox5.Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        For Baris As Integer = 0 To DataGridView1.Rows.Count - 1
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_preparelist_material_log_tbl`(`prepare_list`, `org`, `supply_wo`, `demand_wo`, `pn` , `lot_qty`, `barcode_preparelist`, `barcode_label`, `child`, `result`, `remark`, `created`) VALUES (@prepare_list,@org,@supply_wo,@demand_wo,@pn,@lot_qty,@barcode_preparelist,@barcode_label,@child,@result,@remark,@created)"
            cmd.Parameters.Add("@prepare_list", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(0).Value.ToString
            If CheckBox1.Checked = True Then
                cmd.Parameters.Add("@org", MySqlDbType.VarChar).Value = CheckBox1.Text
            ElseIf CheckBox2.Checked = True Then
                cmd.Parameters.Add("@org", MySqlDbType.VarChar).Value = CheckBox2.Text
            End If
            cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(1).Value.ToString
            cmd.Parameters.Add("@demand_wo", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(2).Value.ToString
            cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(3).Value.ToString
            cmd.Parameters.Add("@lot_qty", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(4).Value.ToString
            cmd.Parameters.Add("@barcode_preparelist", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(5).Value.ToString
            cmd.Parameters.Add("@barcode_label", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(6).Value.ToString
            cmd.Parameters.Add("@child", MySqlDbType.VarChar).Value = ComboBox1.Text
            cmd.Parameters.Add("@result", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(7).Value.ToString
            cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(8).Value.ToString
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = DataGridView1.Rows(Baris).Cells(9).Value.ToString
            cmd.CommandText = str
            cmd.ExecuteNonQuery()
        Next
        MsgBox("Record Successfully Saved!", vbInformation, Application.ProductName)
        Call bersih()
        conn.Dispose()
        conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormProdPlanMaterial.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'If CheckBox1.Checked = False And CheckBox2.Checked = False Then
        'MsgBox("ORG [" & CheckBox1.Text & "] & [" & CheckBox2.Text & "] not checked!", vbCritical)
        'ElseIf CheckBox1.Checked = True Then
        'Call connection()
        'str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox1.Text & "'"
        'cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        'rd = cmd.ExecuteReader
        'rd.Read()
        'Try
        'If Not rd.HasRows Then
        'Label9.Text = "NG"
        'Label9.Font = New Font("Arial", 72)
        'Label9.ForeColor = Color.Red
        'Label11.Text = "Production plan hasn't been updated yet"
        'Label11.Font = New Font("Arial Narrow", 14)
        'Label11.ForeColor = Color.Crimson
        'TextBox1.Text = ""
        'Else
        'Label9.Text = ""
        'Label11.Text = ""
        'TextBox1.Text = rd("supply_wo")
        'TextBox2.Text = rd("demand_wo")
        'TextBox3.Text = rd("pn")
        'TextBox4.Text = rd("lot_qty")
        'TextBox5.Focus()

        'Call LoadBOM34()
        'End If
        'Catch ex As Exception
        'MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Finally
        'rd.Close()
        'End Try
        'ElseIf CheckBox2.Checked = True Then
        'Call connection()
        'str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox2.Text & "'"
        'cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        'rd = cmd.ExecuteReader
        'rd.Read()
        'Try
        'If Not rd.HasRows Then
        'Label9.Text = "NG"
        'Label9.Font = New Font("Arial", 72)
        'Label9.ForeColor = Color.Red
        'Label11.Text = "Production plan hasn't been updated yet"
        'Label11.Font = New Font("Arial Narrow", 14)
        'Label11.ForeColor = Color.Crimson
        'TextBox1.Text = ""
        'Else
        'Label9.Text = ""
        'Label11.Text = ""
        'TextBox1.Text = rd("supply_wo")
        'TextBox2.Text = rd("demand_wo")
        'TextBox3.Text = rd("pn")
        'TextBox4.Text = rd("lot_qty")
        'TextBox5.Focus()

        'Call LoadBOM34()
        'End If
        'Catch ex As Exception
        'MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Finally
        'rd.Close()
        'End Try
        'End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'Call connection()
        'str = "SELECT ebt, supply_wo, wo, wo_qty FROM prodsys_prod_plan_tbl WHERE wo='" & TextBox2.Text & "'"
        'cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        'rd = cmd.ExecuteReader
        'rd.Read()
        'Try
        'If Not rd.HasRows Then
        'Label9.Text = "NG"
        'Label9.Font = New Font("Arial", 72)
        'Label9.ForeColor = Color.Red
        'Label11.Text = "Production plan hasn't been updated yet"
        'Label11.Font = New Font("Arial Narrow", 14)
        'Label11.ForeColor = Color.Crimson
        'TextBox2.Text = ""
        'Else
        'Label9.Text = ""
        'Label11.Text = ""
        'TextBox1.Text = rd("supply_wo")
        'TextBox2.Text = rd("wo")
        'TextBox3.Text = rd("ebt")
        'TextBox4.Text = rd("wo_qty")
        'TextBox5.Focus()

        'Call LoadBOM12()
        'End If
        'Catch ex As Exception
        'MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Finally
        'rd.Close()
        'End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call connection()
        If ComboBox1.SelectedItem = "BOM" Then
            'JIKA SUMBER NYA DARI BOM/QUALITY
            If RadioButton1.Checked = True Then 'BPR
                If CheckBox1.Checked = False And CheckBox2.Checked = False Then
                    MsgBox("ORG [" & CheckBox1.Text & "] & [" & CheckBox2.Text & "] not checked!", vbCritical)
                ElseIf CheckBox1.Checked = True Then
                    Call connection()
                    str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox1.Text & "'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    rd.Read()
                    Try
                        If Not rd.HasRows Then
                            Label9.Text = "NG"
                            Label9.Font = New Font("Arial", 72)
                            Label9.ForeColor = Color.Red
                            Label11.Text = "Production plan hasn't been updated yet"
                            Label11.Font = New Font("Arial Narrow", 14)
                            Label11.ForeColor = Color.Crimson
                            TextBox1.Text = ""
                        Else
                            Label9.Text = ""
                            Label11.Text = ""
                            TextBox1.Text = rd("supply_wo")
                            TextBox2.Text = rd("demand_wo")
                            TextBox3.Text = rd("pn")
                            TextBox4.Text = rd("lot_qty")
                            TextBox5.Focus()

                            Call LoadBOM34()
                        End If
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        rd.Close()
                    End Try
                ElseIf CheckBox2.Checked = True Then
                    Call connection()
                    str = "SELECT pn, supply_wo, demand_wo, lot_qty FROM prodsys2_prod_plan_tbl WHERE supply_wo='" & TextBox1.Text & "' AND org='" & CheckBox2.Text & "'"
                    cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                    rd = cmd.ExecuteReader
                    rd.Read()
                    Try
                        If Not rd.HasRows Then
                            Label9.Text = "NG"
                            Label9.Font = New Font("Arial", 72)
                            Label9.ForeColor = Color.Red
                            Label11.Text = "Production plan hasn't been updated yet"
                            Label11.Font = New Font("Arial Narrow", 14)
                            Label11.ForeColor = Color.Crimson
                            TextBox1.Text = ""
                        Else
                            Label9.Text = ""
                            Label11.Text = ""
                            TextBox1.Text = rd("supply_wo")
                            TextBox2.Text = rd("demand_wo")
                            TextBox3.Text = rd("pn")
                            TextBox4.Text = rd("lot_qty")
                            TextBox5.Focus()

                            Call LoadBOM34()
                        End If
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        rd.Close()
                    End Try
                End If
            ElseIf RadioButton2.Checked = True Then 'DMS
                str = "SELECT ebt, supply_wo, wo, wo_qty FROM prodsys_prod_plan_tbl WHERE wo='" & TextBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If Not rd.HasRows Then
                        Label9.Text = "NG"
                        Label9.Font = New Font("Arial", 72)
                        Label9.ForeColor = Color.Red
                        Label11.Text = "Production plan hasn't been updated yet"
                        Label11.Font = New Font("Arial Narrow", 14)
                        Label11.ForeColor = Color.Crimson
                        TextBox2.Text = ""
                    Else
                        Label9.Text = ""
                        Label11.Text = ""
                        TextBox1.Text = rd("supply_wo")
                        TextBox2.Text = rd("wo")
                        TextBox3.Text = rd("ebt")
                        TextBox4.Text = rd("wo_qty")
                        TextBox5.Focus()

                        Call LoadBOM12()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    rd.Close()
                End Try
            Else
                MsgBox("No Process is selected!", MsgBoxStyle.Information)
                Call bersih()
            End If
        ElseIf ComboBox1.SelectedItem = "MRP" Then
            'JIKA SUMBER NYA DARI MRP/MATERIAL
            If RadioButton1.Checked = True Then 'BPR
                Call LoadMRPBPR()
            ElseIf RadioButton2.Checked = True Then 'DMS
                Call LoadMRPDMS()
            ElseIf RadioButton3.Checked = True Then 'SMT
                Call LoadMRPSMT()
            Else
                MsgBox("No Process is selected!", MsgBoxStyle.Information)
                Call bersih()
            End If
        End If
    End Sub
End Class