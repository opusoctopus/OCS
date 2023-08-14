Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient
Public Class FormRequestICeMMC
    Public Sub StripNonAlphabetCharacters(ByVal input As TextBox)
        ' pattern matches any character that is NOT A-Z (allows upper and lower case alphabets)
        Dim rx As New Regex("[^0-9]")
        If (rx.IsMatch(input.Text)) Then
            Dim startPosition As Integer = input.SelectionStart - 1
            input.Text = rx.Replace(input.Text, "")
            input.SelectionStart = startPosition
        End If
    End Sub

    Sub loaddata()
        rd.Close()
        TextBox4.Text = FormMain.LabelNama.Text

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT supply_wo, pn, qty_req, requested, pn_comp, version, checksum, marking, is_active, created FROM prodsys2_req_ic_emmc_tbl WHERE is_active = '1'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT supply_wo, pn, qty_req, requested, pn_comp, version, checksum, marking, is_active, created FROM prodsys2_req_ic_emmc_tbl WHERE is_active = '0'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)
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

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FormListICeMMC.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormProdPlanICMapping.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Val(TextBox3.Text) >= Val(TextBox9.Text) Then
            MsgBox("REQUESTED QTY MORE THAN STOCK", vbInformation)
        End If

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("CAN'T EMPTY DATA!", vbExclamation)
            Call bersih()
            Call loaddata()
        Else
            Call connection()
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_req_ic_emmc_tbl` (`supply_wo`, `pn`, `qty_req` , `requested`, `pn_comp`, `version`, `checksum`, `marking`, `is_active`, `created`) VALUES (@supply_wo,@pn,@qty_req,@requested,@pn_comp,@version,@checksum,@marking,@is_active,@created)"
            cmd.Parameters.Add("@supply_wo", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@pn", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@qty_req", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@requested", MySqlDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("@pn_comp", MySqlDbType.VarChar).Value = TextBox8.Text
            cmd.Parameters.Add("@version", MySqlDbType.VarChar).Value = TextBox7.Text
            cmd.Parameters.Add("@checksum", MySqlDbType.VarChar).Value = TextBox6.Text
            cmd.Parameters.Add("@marking", MySqlDbType.VarChar).Value = TextBox5.Text
            cmd.Parameters.Add("@is_active", MySqlDbType.VarChar).Value = CheckBox1.Text
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
            cmd.CommandText = str
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Call bersih()
                Call loaddata()
            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call bersih()
    End Sub

    Private Sub FormRequestICeMMC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call loaddata()
        CheckBox1.Checked = True
        CheckBox1.Enabled = False
        TextBox4.ReadOnly = True
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        StripNonAlphabetCharacters(TextBox3)
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        StripNonAlphabetCharacters(TextBox9)
    End Sub
End Class