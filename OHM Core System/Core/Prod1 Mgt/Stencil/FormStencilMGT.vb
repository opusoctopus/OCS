Imports MySql.Data.MySqlClient
Public Class FormStencilMGT

    Dim arrImage1() As Byte
    Dim arrImage2() As Byte

    Private Sub FormStencilMGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadDataInspection()
        Call LoadDataInLine()
    End Sub

    Sub customdgv2()
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
        With DataGridView3 'Ganti dengan nama datagridview kalian
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

    Sub LoadDataInspection()
        Try
            str = "SELECT kode_stencil as KodeStencil, 
                    pn_pcb as PartNumber, 
                    prod_qty as Qty, 
                    cleaning_status as Cleaned, 
                    tension_pa as TensionP1, 
                    tension_pb as TensionP2, 
                    tension_pc as TensionP3, 
                    tension_pd as TensionP4, 
                    tension_pe as TensionP5, 
                    pic as PIC, 
                    DATE_FORMAT(STR_TO_DATE(created,'%d/%m/%Y'), '%Y/%m/%d') as Created
                   FROM prodsys2_stencil_trans_in_tbl ORDER BY Created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInspection()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT line as Line, 
                    kode_stencil as KodeStencil, 
                    pn_pcb as PartNumber,       
                    pic as PIC, 
                    created as Created 
                   FROM prodsys2_stencil_trans_out_tbl WHERE status='AVAILABLE' ORDER BY created DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        rd.Close()
        conn.Close()
        Call BuatKolomInLine()
    End Sub

    Sub BuatKolomInspection()
        DataGridView2.Columns(0).Width = 130
        DataGridView2.Columns(1).Width = 120
        DataGridView2.Columns(2).Width = 80
        DataGridView2.Columns(3).Width = 80
        DataGridView2.Columns(4).Width = 100
        DataGridView2.Columns(5).Width = 100
        DataGridView2.Columns(6).Width = 100
        DataGridView2.Columns(7).Width = 100
        DataGridView2.Columns(8).Width = 100
        DataGridView2.Columns(9).Width = 100
        DataGridView2.Columns(10).Width = 130
    End Sub

    Sub BuatKolomInLine()
        DataGridView3.Columns(0).Width = 80
        DataGridView3.Columns(1).Width = 120
        DataGridView3.Columns(2).Width = 120
        DataGridView3.Columns(3).Width = 80
        DataGridView3.Columns(4).Width = 130
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
        FormStencilEvidence1.ShowDialog()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)
        FormStencilEvidence2.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormStencilInventory.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FormStencilOut.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FormAttachMetalmask.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        FormStencilRack.ShowDialog()
    End Sub
End Class