Imports MySql.Data.MySqlClient

Public Class FormMappingLocator

    Sub BuatKolom()
        'FormMappingFG.DataGridView1.Columns(0).Width = 100 'txnid
        FormMappingFG.DataGridView1.Columns(0).Width = 100 'locator
        FormMappingFG.DataGridView1.Columns(1).Width = 100 'wo
        FormMappingFG.DataGridView1.Columns(2).Width = 120 'pn
        FormMappingFG.DataGridView1.Columns(3).Width = 160 'model
        FormMappingFG.DataGridView1.Columns(4).Width = 100 'demand_date
        FormMappingFG.DataGridView1.Columns(5).Width = 60 'supply_qty
        FormMappingFG.DataGridView1.Columns(6).Width = 60 'inspect_qty
        FormMappingFG.DataGridView1.Columns(7).Width = 70 'remained_qty
        FormMappingFG.DataGridView1.Columns(8).Width = 180 'date_in
        FormMappingFG.DataGridView1.Columns(9).Width = 180 'date_out
        'DataGridView1.Columns(11).Width = 180 'percent
    End Sub

    Sub BuatKolom2()
        'FormMappingFG.DataGridView1.Columns(0).Width = 100 'txnid
        FormMappingFG.DataGridView2.Columns(0).Width = 100 'locator
        FormMappingFG.DataGridView2.Columns(1).Width = 80 'wo
        FormMappingFG.DataGridView2.Columns(2).Width = 100 'pn
        FormMappingFG.DataGridView2.Columns(3).Width = 100 'pst
        FormMappingFG.DataGridView2.Columns(4).Width = 140 'qty
    End Sub

    Sub LoadData()
        Try
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, model as Model, demand_date as PST, supply_qty as Qty, inspect_qty as OQA, remained_qty as Remain, date_in as DateIn, date_out as DateOut FROM prodsys2_whfg_in_status_tbl WHERE locator IS NULL ORDER BY locator ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormMappingFG.DataGridView1.DataSource = dt
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
            FormMappingFG.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom2()
    End Sub

    Sub showcount()
        FormMappingFG.Label22.Text = FormMappingFG.DataGridView1.RowCount
        FormMappingFG.Label2.Text = FormMappingFG.DataGridView2.RowCount
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_whfg_monitoring_tbl`(`locator`, `wo`, `pn` , `model`, `demand_date`, `qty`, `date_in`) VALUES (@locator,@wo,@pn,@model,@demand_date,@qty,@date_in)"
        cmd.Parameters.Add("@locator", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@model", MySqlDbType.VarChar).Value = TextBox4.Text
        cmd.Parameters.Add("@demand_date", MySqlDbType.VarChar).Value = DateTimePicker1.Text
        cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@date_in", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
            rd.Close()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try

        Call connection()
        Dim sql As String
        sql = "UPDATE prodsys2_whfg_in_status_tbl SET locator = @locator2 WHERE wo = @wo2"
        cmd.Connection = conn
        cmd.CommandText = sql
        cmd.Parameters.AddWithValue("@wo2", TextBox2.Text)
        cmd.Parameters.AddWithValue("@locator2", ComboBox1.Text)
        Dim r As Integer
        r = cmd.ExecuteNonQuery()
        cmd.Parameters.Clear()
        Call LoadData()
        Call LoadData2()
        Call showcount()
        TextBox2.Select()
        Me.Close()
    End Sub
End Class