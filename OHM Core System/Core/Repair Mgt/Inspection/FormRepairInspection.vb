Imports MySql.Data.MySqlClient

Public Class FormRepairInspection

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
            .ColumnHeadersHeight = 30
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
            .BorderStyle = BorderStyle.None
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersDefaultCellStyle.WrapMode = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 30
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

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 50 'TXN_ID '4
        DataGridView1.Columns(1).Width = 80 'SN '5
        DataGridView1.Columns(2).Width = 100 'WO '6
        DataGridView1.Columns(3).Width = 100 'PN '7
        DataGridView1.Columns(4).Width = 150 'NG '8
        DataGridView1.Columns(5).Width = 60 'STATUS '9
        DataGridView1.Columns(6).Width = 110 'REPAIRMAN '10
        DataGridView1.Columns(7).Width = 80 'APPROVED '11
        DataGridView1.Columns(8).Width = 120 'APPROVAL '12
        DataGridView1.Columns(9).Width = 80 'APPROVED LQC '13
        DataGridView1.Columns(10).Width = 120 'APPROVAL LQC '14
        DataGridView1.Columns(11).Width = 100 'TIMESTAMPS '15
    End Sub

    Sub BuatKolomSUM()
        DataGridView2.Columns(0).Width = 80 'SN
        DataGridView2.Columns(1).Width = 100 'PROGRESS
        DataGridView2.Columns(2).Width = 80 'APP LEADER
        DataGridView2.Columns(3).Width = 80 'APP LQC
        DataGridView2.Columns(4).Width = 60 'RESULT
    End Sub

    Sub Bersih()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        PictureBox1.Image = Nothing

        TextBox2.Select()
    End Sub

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
        Label1.Text = DataGridView2.RowCount
    End Sub

    Sub LoadData()
        Try
            str = "SELECT txn_id as TxnID, sn as BarCode, supply_wo as WO, pn as PartNumber, ng as NG, status as Status, repairman as Repairman, approved as VLeader, approval as Leader, approved_lqc as VLqc, approval_lqc as LQC, created as Timestamps FROM prodsys2_repair_inspection_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()

        Dim columnbutton1 As New DataGridViewButtonColumn
        columnbutton1.Width = 60
        columnbutton1.Name = "columnbutton1"
        columnbutton1.HeaderText = "Leader"
        columnbutton1.Text = "Verify"
        columnbutton1.UseColumnTextForButtonValue = True
        columnbutton1.FlatStyle = FlatStyle.Popup
        columnbutton1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(0, columnbutton1)

        Dim columnbutton4 As New DataGridViewButtonColumn
        columnbutton4.Width = 60
        columnbutton4.Name = "columnbutton4"
        columnbutton4.HeaderText = "LQC"
        columnbutton4.Text = "Verify"
        columnbutton4.UseColumnTextForButtonValue = True
        columnbutton4.FlatStyle = FlatStyle.Popup
        columnbutton4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(1, columnbutton4)

        Dim columnbutton2 As New DataGridViewButtonColumn
        columnbutton2.Width = 90
        columnbutton2.Name = "columnbutton2"
        columnbutton2.HeaderText = "Repairman"
        columnbutton2.Text = "Update"
        columnbutton2.UseColumnTextForButtonValue = True
        columnbutton2.FlatStyle = FlatStyle.Popup
        columnbutton2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(2, columnbutton2)

        Dim columnbutton3 As New DataGridViewButtonColumn
        columnbutton3.Width = 60
        columnbutton3.Name = "columnbutton3"
        columnbutton3.HeaderText = "Detail"
        columnbutton3.Text = "Detail"
        columnbutton3.UseColumnTextForButtonValue = True
        columnbutton3.FlatStyle = FlatStyle.Popup
        columnbutton3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Insert(3, columnbutton3)

    End Sub

    Sub LoadDataSum()
        Try
            str = "SELECT sn as SN, `status` AS `PROGRESS`, approved as LEADER, approved_lqc as LQC,
                       CASE
                           WHEN `status` = 'OPEN' AND approved = 'NO' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'NO' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'YES' AND approved_lqc = 'NO' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'NO' AND approved_lqc = 'YES' THEN 'OPEN'
					                 WHEN `status` = 'CLOSED' AND approved = 'YES' AND approved_lqc = 'YES' THEN 'CLOSED'
                           ELSE 'Unknown'
                       END AS result
                   FROM prodsys2_repair_inspection_tbl;"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomSUM()
    End Sub

    Sub LoadDataPencarian()
        Try
            str = "SELECT txn_id as TxnID, sn as BarCode, supply_wo as WO, pn as PartNumber, ng as NG, status as Status, repairman as Repairman, approved as VLeader, approval as Leader, approved_lqc as VLqc, approval_lqc as LQC, created as Timestamps FROM prodsys2_repair_inspection_tbl WHERE sn LIKE '" & TextBox1.Text & "'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
        Call showcount()
    End Sub

    Private Sub FormRepairInspection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call LoadDataSum()
        Call showcount()

        TextBox6.Text = FormMain.LabelNama.Text
        'TextBox6.Enabled = False

        TextBox2.Select()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox1.Text.StartsWith("IO")) Then
                MsgBox("SN [" & TextBox1.Text & "] INVALID CHARACTER", MsgBoxStyle.Information)
                Call customdgv2()
                Call LoadData()
                Call showcount()
                TextBox1.SelectAll()
            Else
                Call customdgv2()
                Call LoadDataPencarian()
                Call showcount()
                TextBox1.SelectAll()
            End If
        End If
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.SelectAll()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportDgvToExcel(Me.DataGridView1)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If Not (TextBox2.Text.StartsWith("IO")) Then
                MsgBox("SN [" & TextBox2.Text & "] INVALID CHARACTER", MsgBoxStyle.Information)
                TextBox2.SelectAll()
            Else
                Call connection()
                str = "SELECT supply_wo, pn FROM prodsys_sn_tbl WHERE sn='" & TextBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        TextBox3.Text = rd.Item("supply_wo")
                        TextBox4.Text = rd.Item("pn")

                        e.Handled = True
                        SendKeys.SendWait("{TAB}")
                        TextBox5.Select()
                    Else
                        e.Handled = True
                        SendKeys.SendWait("{TAB}")
                        TextBox3.Select()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try
            End If
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        FormEvidenceRepairBefore.ShowDialog()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        FormEvidenceRepairAfter.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim sqlQuery As String = "INSERT INTO prodsys2_repair_inspection_tbl (`sn`, `supply_wo`, `pn`, `ng`, `status`, `before_inspect`, `repairman`, `approved`, `approved_lqc`, `created`) VALUES (@sn,@supply_wo,@pn,@ng,@status,@before_inspect,@repairman,@approved,@approved_lqc,@created)"
            Dim sqlcom As New MySqlCommand(sqlQuery, conn)
            Dim ms As New IO.MemoryStream()
            PictureBox1.Image.Save(ms, Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New MySqlParameter("@before_inspect", MySqlDbType.LongBlob)
            p.Value = data
            sqlcom.Parameters.Add(p)
            sqlcom.Parameters.AddWithValue("@sn", TextBox2.Text)
            sqlcom.Parameters.AddWithValue("@supply_wo", TextBox3.Text)
            sqlcom.Parameters.AddWithValue("@pn", TextBox4.Text)
            sqlcom.Parameters.AddWithValue("@ng", TextBox5.Text)
            sqlcom.Parameters.AddWithValue("@status", "OPEN")
            sqlcom.Parameters.AddWithValue("@repairman", TextBox6.Text)
            sqlcom.Parameters.AddWithValue("@approved", "NO")
            sqlcom.Parameters.AddWithValue("@approved_lqc", "NO")
            sqlcom.Parameters.AddWithValue("@created", System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
            conn.Open()
            sqlcom.ExecuteNonQuery()
            MsgBox("Data berhasil disimpan!", MsgBoxStyle.Information)
            Call Bersih()

            Try
                str = "SELECT txn_id as TxnID, sn as BarCode, supply_wo as WO, pn as PartNumber, ng as NG, status as Status, repairman as Repairman, approved as VLeader, approval as Leader, approved_lqc as VLqc, approval_lqc as LQC, created as Timestamps FROM prodsys2_repair_inspection_tbl"
                Dim da As New MySqlDataAdapter(str, conn)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                conn.Close()
            End Try
            Call BuatKolom()
            Call showcount()
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Leader" Then
            FormRepairInspectionVerify.Label8.Text = DataGridView1.SelectedCells(4).Value.ToString 'TXN_ID
            FormRepairInspectionVerify.Label4.Text = DataGridView1.SelectedCells(5).Value.ToString 'SN
            FormRepairInspectionVerify.Label5.Text = DataGridView1.SelectedCells(6).Value.ToString 'WO
            FormRepairInspectionVerify.Label6.Text = DataGridView1.SelectedCells(7).Value.ToString 'PN
            FormRepairInspectionVerify.Label7.Text = DataGridView1.SelectedCells(8).Value.ToString 'NG
            FormRepairInspectionVerify.Label9.Text = DataGridView1.SelectedCells(9).Value.ToString 'STATUS

            FormRepairInspectionVerify.ShowDialog()
        ElseIf DataGridView1.Columns(e.ColumnIndex).HeaderText = "Repairman" Then
            FormRepairInspectionUpdate.Label8.Text = DataGridView1.SelectedCells(4).Value.ToString 'TXN_ID
            FormRepairInspectionUpdate.Label4.Text = DataGridView1.SelectedCells(5).Value.ToString 'SN
            FormRepairInspectionUpdate.Label5.Text = DataGridView1.SelectedCells(6).Value.ToString 'WO
            FormRepairInspectionUpdate.Label6.Text = DataGridView1.SelectedCells(7).Value.ToString 'PN
            FormRepairInspectionUpdate.Label7.Text = DataGridView1.SelectedCells(8).Value.ToString 'NG
            FormRepairInspectionUpdate.Label9.Text = DataGridView1.SelectedCells(9).Value.ToString 'STATUS

            FormRepairInspectionUpdate.ShowDialog()
        ElseIf DataGridView1.Columns(e.ColumnIndex).HeaderText = "Detail" Then
            FormRepairInspectionDetail.Label8.Text = DataGridView1.SelectedCells(4).Value.ToString 'TXN_ID
            FormRepairInspectionDetail.Label4.Text = DataGridView1.SelectedCells(5).Value.ToString 'SN
            FormRepairInspectionDetail.Label5.Text = DataGridView1.SelectedCells(6).Value.ToString 'WO
            FormRepairInspectionDetail.Label6.Text = DataGridView1.SelectedCells(7).Value.ToString 'PN
            FormRepairInspectionDetail.Label7.Text = DataGridView1.SelectedCells(8).Value.ToString 'NG
            FormRepairInspectionDetail.Label9.Text = DataGridView1.SelectedCells(9).Value.ToString 'STATUS
            FormRepairInspectionDetail.Label13.Text = DataGridView1.SelectedCells(11).Value.ToString 'STATUS
            FormRepairInspectionDetail.Label14.Text = DataGridView1.SelectedCells(13).Value.ToString 'STATUS

            FormRepairInspectionDetail.ShowDialog()
        ElseIf DataGridView1.Columns(e.ColumnIndex).HeaderText = "LQC" Then
            FormRepairInspectionLQC.Label8.Text = DataGridView1.SelectedCells(4).Value.ToString 'TXN_ID
            FormRepairInspectionLQC.Label4.Text = DataGridView1.SelectedCells(5).Value.ToString 'SN
            FormRepairInspectionLQC.Label5.Text = DataGridView1.SelectedCells(6).Value.ToString 'WO
            FormRepairInspectionLQC.Label6.Text = DataGridView1.SelectedCells(7).Value.ToString 'PN
            FormRepairInspectionLQC.Label7.Text = DataGridView1.SelectedCells(8).Value.ToString 'NG
            FormRepairInspectionLQC.Label9.Text = DataGridView1.SelectedCells(9).Value.ToString 'STATUS

            FormRepairInspectionLQC.ShowDialog()
        End If
    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick
        TextBox2.SelectAll()
    End Sub
End Class