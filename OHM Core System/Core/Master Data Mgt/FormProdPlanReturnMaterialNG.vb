Imports MySql.Data.MySqlClient

Public Class FormProdPlanReturnMaterialNG
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

    Sub LoadData()
        Try
            str = "select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try

        Try
            str = "select ebt, supply_wo, supply_qty, wo, assembly_partno, demand_line, wo_qty, created from prodsys_prod_plan_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        conn.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Me.TabControl1.SelectedIndex = 0 Then ' SMT PLAN
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
                conn.Close()
            ElseIf Me.TabControl1.SelectedIndex = 1 Then ' ASP
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
                conn.Close()
            End If
        End If
    End Sub

    Private Sub FormProdPlanReturnMaterialNG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormReturnMaterialNG.TextBox3.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' pn
        FormReturnMaterialNG.TextBox2.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' wo supply
        Me.Close()
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        FormReturnMaterialNG.TextBox3.Text = Me.DataGridView2.SelectedCells(0).Value.ToString ' pn
        FormReturnMaterialNG.TextBox2.Text = Me.DataGridView2.SelectedCells(3).Value.ToString ' wo demand
        Me.Close()
    End Sub
End Class