Public Class FormApprovalReqICeMMC
    Sub LoadData()
        rd.Close()
        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT supply_wo, pn, qty_req, requested, pn_comp, version, checksum, marking, is_active, created FROM prodsys2_req_ic_emmc_tbl WHERE is_active = '1'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)

        da = New MySql.Data.MySqlClient.MySqlDataAdapter("SELECT supply_wo, pn, qty_req, requested, pn_comp, version, checksum, marking, is_active, created FROM prodsys2_req_ic_emmc_tbl WHERE is_active = '0'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(112, 122, 131)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
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
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(112, 122, 131)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(112, 122, 131)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 22
            End With
        End With
    End Sub

    Private Sub FormApprovalReqICeMMC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call LoadData()
        Call customdgv2()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormPopupApprovalReqICeMMC.TextBox1.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' SUPPLY WO
        FormPopupApprovalReqICeMMC.TextBox2.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' EBT
        FormPopupApprovalReqICeMMC.TextBox3.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' QTY REQ
        FormPopupApprovalReqICeMMC.TextBox4.Text = Me.DataGridView1.SelectedCells(3).Value.ToString ' REQUESTED
        FormPopupApprovalReqICeMMC.TextBox5.Text = Me.DataGridView1.SelectedCells(4).Value.ToString ' PN COMPONENT
        FormPopupApprovalReqICeMMC.TextBox6.Text = Me.DataGridView1.SelectedCells(5).Value.ToString ' SW VERSION
        FormPopupApprovalReqICeMMC.TextBox7.Text = Me.DataGridView1.SelectedCells(6).Value.ToString ' CHECKSUM
        FormPopupApprovalReqICeMMC.TextBox8.Text = Me.DataGridView1.SelectedCells(7).Value.ToString ' MARKING

        FormPopupApprovalReqICeMMC.TextBox1.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox2.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox3.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox4.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox5.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox6.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox7.ReadOnly = True
        FormPopupApprovalReqICeMMC.TextBox8.ReadOnly = True

        Call connection()
        str = "SELECT * FROM prodsys2_req_ic_emmc_tbl WHERE supply_wo='" & DataGridView1.SelectedCells(0).Value.ToString & "'"
        cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
        rd = cmd.ExecuteReader
        rd.Read()
        Try
            If rd.HasRows Then
                FormPopupApprovalReqICeMMC.TextBox9.Text = rd("id")
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            rd.Close()
            conn.Close()
            conn.Dispose()
        End Try

        FormPopupApprovalReqICeMMC.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call LoadData()
    End Sub
End Class