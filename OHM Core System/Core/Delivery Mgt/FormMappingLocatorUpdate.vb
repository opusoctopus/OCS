Imports MySql.Data.MySqlClient

Public Class FormMappingLocatorUpdate

    Sub BuatKolom()
        FormMappingFG.DataGridView2.Columns(0).Width = 100 'locator
        FormMappingFG.DataGridView2.Columns(1).Width = 80 'wo
        FormMappingFG.DataGridView2.Columns(2).Width = 100 'pn
        FormMappingFG.DataGridView2.Columns(3).Width = 100 'pst
        FormMappingFG.DataGridView2.Columns(4).Width = 140 'qty
    End Sub

    Sub BuatKolomDashboard()
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

    Sub LoadDataDashboard()
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
        Call BuatKolomDashboard()
    End Sub

    Sub LoadData()
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
        Call BuatKolom()
    End Sub

    Sub showcountdashboard()
        FormMappingFG.Label22.Text = FormMappingFG.DataGridView1.RowCount
        FormMappingFG.Label2.Text = FormMappingFG.DataGridView2.RowCount
    End Sub

    Private Sub FormMappingLocatorUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        rd.Close()
        Call connection()
        str = "SELECT * FROM prodsys2_whfg_monitoring_tbl WHERE wo='" & TextBox2.Text & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                rd.Close()
                Dim sql As String
                sql = "UPDATE prodsys2_whfg_monitoring_tbl SET locator = @locator, qty=@qty, demand_date=@demand_date WHERE wo = @wo"
                cmd.Connection = conn
                cmd.CommandText = sql
                cmd.Parameters.AddWithValue("@wo", TextBox2.Text)
                cmd.Parameters.AddWithValue("@locator", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@qty", textbox1.Text)
                cmd.Parameters.AddWithValue("@demand_date", DateTimePicker1.Text)
                Dim r As Integer
                r = cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
                conn.Close()
                conn.Dispose()
                Call LoadData()
                Call LoadDataDashboard()
                Call showcountdashboard()
                Me.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class