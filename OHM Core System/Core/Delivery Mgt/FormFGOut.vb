Imports MySql.Data.MySqlClient

Public Class FormFGOut
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.SelectedItem = Nothing
    End Sub

    Sub BuatKolom()
        'DataGridView1.Columns(0).Width = 170 'txnid
        FormLocatorFG.DataGridView1.Columns(0).Width = 160 'locator
        FormLocatorFG.DataGridView1.Columns(1).Width = 170 'wo
        FormLocatorFG.DataGridView1.Columns(2).Width = 210 'pn
        FormLocatorFG.DataGridView1.Columns(3).Width = 350 'model
        FormLocatorFG.DataGridView1.Columns(4).Width = 170 'demand_date
        FormLocatorFG.DataGridView1.Columns(5).Width = 100 'supply_qty
        FormLocatorFG.DataGridView1.Columns(6).Width = 100 'inspect_qty
        FormLocatorFG.DataGridView1.Columns(7).Width = 130 'remained_qty
        FormLocatorFG.DataGridView1.Columns(8).Width = 265 'date_in
        FormLocatorFG.DataGridView1.Columns(9).Width = 265 'date_out
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
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, model as Model, demand_date as PST, supply_qty as Qty, inspect_qty as OQA, remained_qty as Remain, date_in as DateIn, date_out as DateOut FROM prodsys2_whfg_in_status_tbl WHERE locator IS NOT NULL AND date_out IS NULL ORDER BY locator ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormLocatorFG.DataGridView1.DataSource = dt
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
        Call BuatKolom()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT * FROM prodsys2_whfg_in_status_tbl WHERE wo='" & TextBox2.Text & "' AND date_out IS NULL"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If Not rd.HasRows Then
                    MsgBox("WO ini tidak ada di warehose atau sudah close delivery", vbExclamation)
                    TextBox2.Text = ""
                    TextBox2.Select()
                    Call bersih()
                Else
                    'TextBox1.Text = rd("txn_id")
                    TextBox3.Text = rd("pn")
                    TextBox4.Text = rd("model")
                    DateTimePicker1.Text = rd("demand_date")
                    ComboBox1.Text = rd("locator")
                    conn.Close()

                    conn.Open()
                    Dim sql As String
                    sql = "UPDATE prodsys2_whfg_in_status_tbl SET date_out = @date_out WHERE wo = @wo"
                    cmd.Connection = conn
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@wo", TextBox2.Text)
                    cmd.Parameters.AddWithValue("@date_out", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
                    Dim r As Integer
                    r = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()

                    Call LoadData()
                    Call LoadData2()
                    Call bersih()

                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
                conn.Dispose()
            Finally
                rd.Close()
                conn.Close()
                conn.Dispose()
            End Try
        End If
    End Sub

    Private Sub FormFGOut_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TextBox1.ReadOnly = True
        TextBox2.Select()
        Call bersih()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)

    End Sub
End Class