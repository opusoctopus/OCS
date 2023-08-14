Imports MySql.Data.MySqlClient

Public Class FormNozzleInventory

    Sub Bersih()
        TextBox2.Clear()
        ComboBox1.SelectedItem = Nothing
        ComboBox1.Text = ""
    End Sub

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

    Sub LoadData()
        Try
            str = "SELECT txn_id as No, nozzle_type as TypeNozzle, qty as QTY FROM prodsys2_nozzle_inventory_tbl"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolom()
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

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 150
    End Sub

    Sub ShowCount()
        Label5.Text = DataGridView1.RowCount
    End Sub

    Private Sub FormNozzleInventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call LoadDataNozzleType()
        Call ShowCount()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.Text = String.Empty Or TextBox2.Text = String.Empty Then
            MsgBox("Data belum lengkap!", MsgBoxStyle.Information)
        Else
            Call connection()
            str = "SELECT * FROM prodsys2_nozzle_inventory_tbl WHERE nozzle_type='" & ComboBox1.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            Try
                If rd.HasRows Then
                    While rd.Read
                        Dim cmd As New MySqlCommand("UPDATE prodsys2_nozzle_inventory_tbl SET qty = @qty WHERE nozzle_type = @nozzle_type", conn)
                        cmd.Parameters.Add("@nozzle_type", MySqlDbType.VarChar).Value = ComboBox1.Text
                        cmd.Parameters.Add("@qty", MySqlDbType.Int64).Value = Val(rd("qty")) + Val(TextBox2.Text)
                        rd.Close()
                        cmd.ExecuteNonQuery()
                        Call LoadData()
                        Call LoadDataNozzleType()
                        Call Bersih()
                    End While
                Else
                    rd.Close()
                    cmd = New MySql.Data.MySqlClient.MySqlCommand
                    cmd.Connection = conn
                    str = "INSERT INTO `prodsys2_nozzle_inventory_tbl` (`nozzle_type`, `qty`, `created`) VALUES (@nozzle_type,@qty,@created)"
                    cmd.Parameters.Add("@nozzle_type", MySqlDbType.VarChar).Value = ComboBox1.Text
                    cmd.Parameters.Add("@qty", MySqlDbType.Int64).Value = TextBox2.Text
                    cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
                    cmd.CommandText = str
                    Try
                        cmd.ExecuteNonQuery()
                        rd.Close()
                        Call Bersih()
                        Call LoadData()
                        Call LoadDataNozzleType()
                    Catch ex As Exception
                        MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End Try
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox2.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Static counter As Integer = 0
        Console.WriteLine("DataGridView1_CellFormatting " & counter)
        counter += 1

        Dim row = DataGridView1.Rows(e.RowIndex)
        If Not row.IsNewRow AndAlso e.ColumnIndex <> 0 Then
            Dim cell = row.Cells(0)
            cell.Value = (e.RowIndex + 1).ToString()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Bersih()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox2.Select()
    End Sub
End Class