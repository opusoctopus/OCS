Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Public Class FormICMapping
    Public Sub StripNonAlphabetCharacters(ByVal input As TextBox)
        ' pattern matches any character that is NOT A-Z (allows upper and lower case alphabets)
        Dim rx As New Regex("[^0-9]")
        If (rx.IsMatch(input.Text)) Then
            Dim startPosition As Integer = input.SelectionStart - 1
            input.Text = rx.Replace(input.Text, "")
            input.SelectionStart = startPosition
        End If
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
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        TextBox1.Focus()
    End Sub

    Sub onn()
        GroupBox1.Enabled = True
        TextBox1.Focus()
    End Sub

    Sub off()
        GroupBox1.Enabled = False
    End Sub

    Sub buatkolom()
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).Width = 50
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).Width = 200
    End Sub

    Sub loaddata()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT pn_comp, version, checksum, marking, qty, worker, created FROM prodsys2_ic_mapping_tbl", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        Call buatkolom()
    End Sub
    Private Sub FormICMapping_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call bersih()
        Call off()
        Call loaddata()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call onn()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("CAN'T EMPTY DATA!", vbExclamation)
            Call bersih()
            Call loaddata()
        Else
            Call connection()
            cmd = New MySql.Data.MySqlClient.MySqlCommand
            cmd.Connection = conn
            str = "INSERT INTO `prodsys2_ic_mapping_tbl` (`pn_comp`, `version`, `checksum` , `marking`, `qty`, `worker`, `created`) VALUES (@pn_comp,@version,@checksum,@marking,@qty,@worker,@created)"
            cmd.Parameters.Add("@pn_comp", MySqlDbType.VarChar).Value = TextBox1.Text
            cmd.Parameters.Add("@version", MySqlDbType.VarChar).Value = TextBox2.Text
            cmd.Parameters.Add("@checksum", MySqlDbType.VarChar).Value = TextBox3.Text
            cmd.Parameters.Add("@marking", MySqlDbType.VarChar).Value = TextBox4.Text
            cmd.Parameters.Add("@qty", MySqlDbType.VarChar).Value = TextBox5.Text
            cmd.Parameters.Add("@worker", MySqlDbType.VarChar).Value = FormMain.LabelNama.Text
            cmd.Parameters.Add("@created", MySqlDbType.VarChar).Value = System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss")
            cmd.CommandText = str
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Call bersih()
                Call loaddata()
                conn.Dispose()
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        StripNonAlphabetCharacters(TextBox5)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class