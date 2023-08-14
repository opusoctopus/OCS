Imports MySql.Data.MySqlClient

Public Class FormNozzleOut

    Sub Bersih()
        TextBox1.Clear()
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        ComboBox3.SelectedItem = Nothing
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID,
                    nozzle_type as NozzleType, 
                    status_inspection as Judgment,
                    remark as Remark,
                    pic as PIC,
                    created as Created
                   FROM prodsys2_nozzle_inspection_tbl ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormNozzleMGT.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolom()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT line as Line,
                    machine as Machine,   
                    nozzle_type as NozzleType,
                    count(nozzle_type) as QTY
                   FROM prodsys2_nozzle_mgt_log_tbl ORDER BY line ASC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormNozzleMGT.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInLine()
    End Sub

    Sub BuatKolom()
        FormNozzleMGT.DataGridView1.Columns(0).Width = 60
        FormNozzleMGT.DataGridView1.Columns(1).Width = 120
        FormNozzleMGT.DataGridView1.Columns(2).Width = 130
        FormNozzleMGT.DataGridView1.Columns(3).Width = 120
        FormNozzleMGT.DataGridView1.Columns(4).Width = 150
        FormNozzleMGT.DataGridView1.Columns(5).Width = 350
    End Sub

    Sub BuatKolomInLine()
        FormNozzleMGT.DataGridView2.Columns(0).Width = 120
        FormNozzleMGT.DataGridView2.Columns(1).Width = 130
        FormNozzleMGT.DataGridView2.Columns(2).Width = 120
        FormNozzleMGT.DataGridView2.Columns(3).Width = 380
    End Sub

    Sub LoadDataNozzleType()
        Dim StrSql As String = "SELECT distinct nozzle_type FROM prodsys2_nozzle_inventory_tbl"
        Dim cmd As New MySqlCommand(StrSql, conn)
        Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
        Dim dt As New DataTable("prodsys2_nozzle_inventory_tbl")
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            ComboBox1.DataSource = dt
            ComboBox1.DisplayMember = "nozzle_type"
            ComboBox1.ValueMember = "nozzle_type"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connection()
        cmd = New MySql.Data.MySqlClient.MySqlCommand
        cmd.Connection = conn
        str = "INSERT INTO `prodsys2_nozzle_mgt_log_tbl` (`nozzle_type`, `machine`, `line`, `pic`, `remark`, `created`)
VALUES (@nozzle_type,@machine,@line,@pic,@remark,@created)"
        cmd.Parameters.Add("@nozzle_type", MySqlDbType.VarChar).Value = ComboBox1.Text
        cmd.Parameters.Add("@machine", MySqlDbType.VarChar).Value = ComboBox2.Text
        cmd.Parameters.Add("@line", MySqlDbType.VarChar).Value = ComboBox3.Text
        cmd.Parameters.Add("@pic", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
        cmd.Parameters.Add("@remark", MySqlDbType.VarChar).Value = TextBox1.Text
        cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
        cmd.CommandText = str
        Try
            cmd.ExecuteNonQuery()
            Call LoadData()
            Call LoadDataInLine()
            Call Bersih()
            rd.Close()
            conn.Dispose()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub FormNozzleOut_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call LoadDataNozzleType()
    End Sub
End Class