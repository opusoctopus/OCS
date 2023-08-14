Imports MySql.Data.MySqlClient
Public Class FormRepairWOList
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 156, 167)
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
            .ColumnHeadersHeight = 35
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 156, 167)
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

    Sub LoadData()
        Try
            str = "SELECT line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created FROM prodsys2_prod_plan_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try

        Try
            str = "SELECT ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created FROM prodsys_prod_plan_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Me.TabControl1.SelectedIndex = 0 Then ' SMT PLAN
                If ComboBox1.Text = "Supply WO" Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    Try
                        str = "select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where supply_wo like '" & TextBox1.Text & "'"
                        Dim da As New MySqlDataAdapter(str, conn)
                        Dim dt As DataTable = New DataTable
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        conn.Close()
                    End Try
                ElseIf ComboBox1.Text = "Partnumber" Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    Try
                        str = "select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where pn like '" & TextBox1.Text & "'"
                        Dim da As New MySqlDataAdapter(str, conn)
                        Dim dt As DataTable = New DataTable
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        conn.Close()
                    End Try
                End If
            ElseIf Me.TabControl1.SelectedIndex = 1 Then ' ASP
                If ComboBox1.Text = "Demand WO" Then
                    DataGridView2.DataSource = Nothing
                    DataGridView2.Refresh()
                    Try
                        str = "select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where wo like '" & TextBox1.Text & "'"
                        Dim da As New MySqlDataAdapter(str, conn)
                        Dim dt As DataTable = New DataTable
                        da.Fill(dt)
                        DataGridView2.DataSource = dt
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        conn.Close()
                    End Try
                ElseIf ComboBox1.Text = "Partnumber" Then
                    DataGridView2.DataSource = Nothing
                    DataGridView2.Refresh()
                    Try
                        str = "select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl where ebt like '" & TextBox1.Text & "'"
                        Dim da As New MySqlDataAdapter(str, conn)
                        Dim dt As DataTable = New DataTable
                        da.Fill(dt)
                        DataGridView2.DataSource = dt
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        conn.Close()
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub FormRepairWOList_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call LoadData()
        Call customdgv2()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormRepairInput.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' wo supply
        FormRepairInput.TextBox2.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' pn
        FormRepairInput.TextBox3.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' model
        Me.Close()
    End Sub
End Class