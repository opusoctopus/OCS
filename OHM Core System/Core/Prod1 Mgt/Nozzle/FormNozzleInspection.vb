Imports MySql.Data.MySqlClient

Public Class FormNozzleInspection

    Dim imgpath As String
    Dim arrImage() As Byte
    Dim sql As String
    Dim result As Integer

    Sub Bersih()
        ComboBox1.Text = ""
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        PictureBox1.Image = Nothing
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID,
                    nozzle_type as NozzleType, 
                    status_inspection as Judgment,
                    remark as Remark,
                    pic as PIC,
                    approved as Approved,
                    approval as Signed,
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
        FormNozzleMGT.DataGridView1.Columns(1).Width = 110
        FormNozzleMGT.DataGridView1.Columns(2).Width = 90
        FormNozzleMGT.DataGridView1.Columns(3).Width = 120
        FormNozzleMGT.DataGridView1.Columns(4).Width = 120
        FormNozzleMGT.DataGridView1.Columns(5).Width = 100
        FormNozzleMGT.DataGridView1.Columns(6).Width = 120
        FormNozzleMGT.DataGridView1.Columns(7).Width = 180
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

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        FormNozzleCam.ShowDialog()
    End Sub

    Private Sub FormNozzleInspection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call LoadDataNozzleType()
        Call Bersih()
        TextBox3.ReadOnly = True
        TextBox3.Text = FormMain.LabelNama.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim sqlQuery As String = "INSERT INTO prodsys2_nozzle_inspection_tbl (`nozzle_type`, `status_inspection`, `remark`, `pic`, `evidence`, `approved`, `created`) VALUES (@nozzle_type,@status_inspection,@remark,@pic,@evidence,@approved,@created)"
            Dim sqlcom As New MySqlCommand(sqlQuery, conn)
            Dim ms As New IO.MemoryStream()
            PictureBox1.Image.Save(ms, Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New MySqlParameter("@evidence", MySqlDbType.LongBlob)
            p.Value = data
            sqlcom.Parameters.Add(p)
            sqlcom.Parameters.AddWithValue("@nozzle_type", ComboBox1.Text)
            sqlcom.Parameters.AddWithValue("@status_inspection", TextBox2.Text)
            sqlcom.Parameters.AddWithValue("@remark", TextBox1.Text)
            sqlcom.Parameters.AddWithValue("@pic", TextBox3.Text)
            sqlcom.Parameters.AddWithValue("@approved", "NO")
            sqlcom.Parameters.AddWithValue("@created", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss"))
            conn.Open()
            sqlcom.ExecuteNonQuery()
            MsgBox("Data berhasil disimpan!", MsgBoxStyle.Information)
            Call Bersih()
            Call LoadData()
            Call LoadDataInLine()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class