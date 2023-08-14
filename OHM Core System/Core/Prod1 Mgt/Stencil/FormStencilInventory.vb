Imports MySql.Data.MySqlClient

Public Class FormStencilInventory

    Dim cmd As MySqlCommand
    Dim dr As MySqlDataReader
    Dim sql As String
    Dim result As Integer
    Private emptyTextBox As Object

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
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(200, 5, 83)
            .DefaultCellStyle.Font = New System.Drawing.Font("Arial", 10.0!)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 24
            End With
        End With
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 180 '
        DataGridView1.Columns(1).Width = 140 '
        DataGridView1.Columns(2).Width = 120 '
        DataGridView1.Columns(3).Width = 120 '
        DataGridView1.Columns(4).Width = 140 '
        DataGridView1.Columns(5).Width = 140 '
        DataGridView1.Columns(6).Width = 100 '
        DataGridView1.Columns(7).Width = 100 '
        DataGridView1.Columns(8).Width = 180 '
    End Sub

    Sub ShowCount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Sub Bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox13.Clear()

        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing

    End Sub

    Sub LoadData()
        Try
            str = "SELECT kode_stencil as CodeStencil, maker_supply as Maker, pn_pcb as PartNumber, layer as Layer, DATE_FORMAT(STR_TO_DATE(pabrication_date,'%d/%m/%Y'), '%Y/%m/%d') as PabricationDate, DATE_FORMAT(STR_TO_DATE(incoming_date,'%d/%m/%Y'), '%Y/%m/%d') as IncomingDate, location as RAK, usages as Usages,
				    CASE
                           WHEN `usages` >=50000 THEN 'EXPIRED'
					       WHEN `usages` >=40000 AND `usages` <50000 THEN 'WARNING'
                           ELSE 'AVAILABLE'
                    END AS Lifetime
                FROM prodsys2_stencil_inventory_tbl ORDER BY IncomingDate DESC;"
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FormStencilSelfCheckIncoming.ShowDialog()
    End Sub

    Private Sub FormStencilInventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox12.Text = "50000"
        TextBox12.Enabled = False
        TextBox13.Select()

        Call customdgv2()
        Call Bersih()
        Call LoadData()
        Call ShowCount()
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If ComboBox3.SelectedItem = Nothing Then
                MsgBox("Kategori pencarian belum dipilih!", MsgBoxStyle.Information)
                TextBox13.SelectAll()
            ElseIf ComboBox3.SelectedItem = "KODE STENCIL" Then
                Try
                    str = "SELECT kode_stencil as CodeStencil, maker_supply as Maker, pn_pcb as PartNumber, layer as Layer, DATE_FORMAT(STR_TO_DATE(pabrication_date,'%d/%m/%Y'), '%Y/%m/%d') as PabricationDate, DATE_FORMAT(STR_TO_DATE(incoming_date,'%d/%m/%Y'), '%Y/%m/%d') as IncomingDate, location as RAK, usages as Usages,
				    CASE
                           WHEN `usages` >=50000 THEN 'EXPIRED'
					       WHEN `usages` >=40000 AND `usages` <50000 THEN 'WARNING'
                           ELSE 'AVAILABLE'
                    END AS Lifetime
                FROM prodsys2_stencil_inventory_tbl WHERE kode_stencil like '" & TextBox13.Text & "';"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView1.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call BuatKolom()
            ElseIf ComboBox3.SelectedItem = "PCB PARTNO" Then
                Try
                    str = "SELECT kode_stencil as CodeStencil, maker_supply as Maker, pn_pcb as PartNumber, layer as Layer, DATE_FORMAT(STR_TO_DATE(pabrication_date,'%d/%m/%Y'), '%Y/%m/%d') as PabricationDate, DATE_FORMAT(STR_TO_DATE(incoming_date,'%d/%m/%Y'), '%Y/%m/%d') as IncomingDate, location as RAK, usages as Usages,
				    CASE
                           WHEN `usages` >=50000 THEN 'EXPIRED'
					       WHEN `usages` >=40000 AND `usages` <50000 THEN 'WARNING'
                           ELSE 'AVAILABLE'
                    END AS Lifetime
                FROM prodsys2_stencil_inventory_tbl WHERE pn_pcb like '" & TextBox13.Text & "';"
                    Dim da As New MySqlDataAdapter(str, conn)
                    Dim dt As DataTable = New DataTable
                    da.Fill(dt)
                    DataGridView1.DataSource = dt
                Catch ex As Exception
                    MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    conn.Close()
                End Try
                Call BuatKolom()
            End If

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Bersih()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ExportDgvToExcel(DataGridView1)
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox2.Select()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox3.Select()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox4.Select()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox5.Select()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox6.Select()
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox7.Select()
        End If
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox8.Select()
        End If
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox9.Select()
        End If
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox10.Select()
        End If
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox11.Select()
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        FormStencilInvPict.TextBox1.Text = DataGridView1.SelectedCells(0).Value.ToString 'kode stencil
        FormStencilInvPict.TextBox2.Text = DataGridView1.SelectedCells(2).Value.ToString 'pn pcb
        FormStencilInvPict.TextBox3.Text = DataGridView1.SelectedCells(5).Value.ToString 'incoming date

        FormStencilInvPict.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Baris As Integer = DataGridView1.Rows.Count - 1 To 0 Step -1
            If DataGridView1.Rows(Baris).Cells(8).Value = "EXPIRED" Then
                DataGridView1.Rows(Baris).Cells(8).Style.BackColor = Color.FromArgb(200, 5, 83)
                DataGridView1.Rows(Baris).Cells(8).Style.ForeColor = Color.White
            End If
        Next
        For Baris As Integer = DataGridView1.Rows.Count - 1 To 0 Step -1
            If DataGridView1.Rows(Baris).Cells(8).Value = "WARNING" Then
                DataGridView1.Rows(Baris).Cells(8).Style.BackColor = Color.Orange
                DataGridView1.Rows(Baris).Cells(8).Style.ForeColor = Color.White
            End If
        Next
        For Baris As Integer = DataGridView1.Rows.Count - 1 To 0 Step -1
            If DataGridView1.Rows(Baris).Cells(8).Value = "AVAILABLE" Then
                DataGridView1.Rows(Baris).Cells(8).Style.BackColor = Color.Green
                DataGridView1.Rows(Baris).Cells(8).Style.ForeColor = Color.White
            End If
        Next
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox13.Select()
    End Sub
End Class