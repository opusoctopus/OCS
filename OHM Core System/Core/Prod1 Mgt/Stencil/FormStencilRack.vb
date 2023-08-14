Imports MySql.Data.MySqlClient

Public Class FormStencilRack

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
    End Sub

    Sub LoadDataTop()
        Try
            str = "SELECT location as Slot,
                    kode_stencil as KodeStencil,
                    pn_pcb as PartNumber,
                    status as Status
                   FROM prodsys2_stencil_rack WHERE layer ='TOP'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomTOP()
    End Sub

    Sub BuatKolomTOP()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 135
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(3).Width = 100
    End Sub

    Sub LoadDataBottom()
        Try
            str = "SELECT location as Slot,
                    kode_stencil as KodeStencil,
                    pn_pcb as PartNumber,
                    status as Status
                   FROM prodsys2_stencil_rack WHERE layer ='BOTTOM'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call BuatKolomBOTTOM()
    End Sub

    Sub BuatKolomBOTTOM()
        DataGridView2.Columns(0).Width = 60
        DataGridView2.Columns(1).Width = 135
        DataGridView2.Columns(2).Width = 120
        DataGridView1.Columns(3).Width = 100
    End Sub

    Sub ShowCount()
        Label2.Text = DataGridView1.RowCount
        Label4.Text = DataGridView2.RowCount
    End Sub

    Private Sub FormStencilRack_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadDataTop()
        Call LoadDataBottom()
        Call ShowCount()
        TextBox1.Select()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = Nothing Then
            Call LoadDataTop()
            Call LoadDataBottom()
            Call ShowCount()
        Else
            TextBox1.Select()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If ComboBox1.SelectedItem = "TOP" Then
                Try
                    str = "SELECT location as Slot,
                    kode_stencil as KodeStencil,
                    pn_pcb as PartNumber,
                    status as Status
                   FROM prodsys2_stencil_rack WHERE layer ='TOP' AND kode_stencil = '" & TextBox1.Text & "'"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView1.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call BuatKolomTOP()
                Call ShowCount()
            ElseIf ComboBox1.SelectedItem = "BOTTOM" Then
                Try
                    str = "SELECT location as Slot,
                    kode_stencil as KodeStencil,
                    pn_pcb as PartNumber,
                    status as Status
                   FROM prodsys2_stencil_rack WHERE layer ='BOTTOM' AND kode_stencil = '" & TextBox1.Text & "'"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView2.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call BuatKolomBOTTOM()
                Call ShowCount()
            Else
                MsgBox("Layer belum dipilih!", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Select()
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Baris As Integer = DataGridView1.Rows.Count - 1 To 0 Step -1
            If DataGridView1.Rows(Baris).Cells(3).Value = "LINE" Then
                DataGridView1.Rows(Baris).Cells(3).Style.BackColor = Color.FromArgb(200, 5, 83)
                DataGridView1.Rows(Baris).Cells(3).Style.ForeColor = Color.White
            End If
        Next
        For Baris As Integer = DataGridView2.Rows.Count - 1 To 0 Step -1
            If DataGridView2.Rows(Baris).Cells(3).Value = "LINE" Then
                DataGridView2.Rows(Baris).Cells(3).Style.BackColor = Color.FromArgb(200, 5, 83)
                DataGridView2.Rows(Baris).Cells(3).Style.ForeColor = Color.White
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormStencilRakUpdateTOP.ComboBox1.Text = Me.DataGridView1.SelectedCells(0).Value.ToString ' slot rak
        FormStencilRakUpdateTOP.TextBox9.Text = Me.DataGridView1.SelectedCells(1).Value.ToString ' kode stencil
        FormStencilRakUpdateTOP.TextBox1.Text = Me.DataGridView1.SelectedCells(2).Value.ToString ' pcb pn

        FormStencilRakUpdateTOP.ShowDialog()
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.SelectAll()
    End Sub
End Class