Imports MySql.Data.MySqlClient
Public Class FormFGIn

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub

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
        Call showcount()
    End Sub

    Sub showcount()
        FormMappingFG.Label22.Text = FormMappingFG.DataGridView1.RowCount
    End Sub

    Sub NomorOtomatis()
        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM prodsys2_whfg_in_status_tbl WHERE txn_id IN (SELECT MAX(txn_id) FROM prodsys2_whfg_in_status_tbl)", conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            UrutanKode = System.DateTime.Now.ToString("yyyyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            UrutanKode = System.DateTime.Now.ToString("yyyyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        TextBox1.Text = UrutanKode
        rd.Close()
        conn.Close()
    End Sub

    Private Sub FormFGIn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call NomorOtomatis()
        TextBox1.ReadOnly = True
        TextBox2.Select()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim remained, supply, inspect As Integer
        supply = TextBox5.Text
        inspect = TextBox6.Text
        remained = Val(supply) - Val(inspect)
        Label9.Text = remained

        conn.Open()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_whfg_in_status_tbl`(`txn_id`, `wo`, `pn` , `model`, `demand_date`, `supply_qty`, `inspect_qty`, `remained_qty`, `date_in`) VALUES (@txn_id,@wo,@pn,@model,@demand_date,@supply_qty,@inspect_qty,@remained_qty,@date_in)"
        cmd.Parameters.Add("@txn_id", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@wo", MySqlDbType.VarChar).Value = TextBox2.Text
        cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox3.Text
        cmd.Parameters.Add("@model", MySqlDbType.VarChar).Value = TextBox4.Text
        cmd.Parameters.Add("@demand_date", MySqlDbType.VarChar).Value = DateTimePicker1.Text
        cmd.Parameters.Add("@supply_qty", MySqlDbType.VarChar).Value = TextBox5.Text
        cmd.Parameters.Add("@inspect_qty", MySqlDbType.VarChar).Value = TextBox6.Text
        cmd.Parameters.Add("@remained_qty", MySqlDbType.VarChar).Value = Label9.Text
        cmd.Parameters.Add("@date_in", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
            conn.Close()
            Call LoadData()
            Call bersih()
            Call NomorOtomatis()
            TextBox2.Select()
            rd.Close()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            conn.Open()
            str = "SELECT DISTINCT ebt, wo, assembly_partno, supply_qty FROM prodsys_prod_plan_tbl WHERE wo='" & TextBox2.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If Not rd.HasRows Then
                    MsgBox("Prod. Plan belum update, silahkan input manual!", vbExclamation)
                    TextBox3.Select()
                Else
                    TextBox3.Text = rd("ebt")
                    TextBox4.Text = rd("assembly_partno")
                    TextBox5.Text = rd("supply_qty")
                    TextBox6.Select()
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
End Class