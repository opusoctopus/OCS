Public Class FormProdPlanICMapping
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
        FormRequestICeMMC.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' SUPPLY WO
        FormRequestICeMMC.TextBox2.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' PN
        FormRequestICeMMC.TextBox3.Text = Me.DataGridView1.SelectedCells(7).Value.ToString ' QTY
        Me.Close()
        FormRequestICeMMC.TextBox3.Focus()
    End Sub

    Private Sub FormProdPlanICMapping_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call customdgv2()

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            da = New MySql.Data.MySqlClient.MySqlDataAdapter("select line, org, supply_wo, model, pn, demand_wo, pcb, lot_qty, pst, remark, created from prodsys2_prod_plan_tbl where supply_wo like '" & TextBox1.Text & "' ", conn)
            ds = New DataSet
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        End If
    End Sub
End Class