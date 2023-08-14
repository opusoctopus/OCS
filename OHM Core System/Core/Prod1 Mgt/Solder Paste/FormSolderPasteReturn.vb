Imports MySql.Data.MySqlClient

Public Class FormSolderPasteReturn
    Dim sql As String
    Dim result As Integer

    Sub KolomCoolingCase()
        FormSolderPasteMGT.DataGridView1.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView1.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView1.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView1.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView1.Columns(4).Width = 180 'start time
        FormSolderPasteMGT.DataGridView1.Columns(5).Width = 180 'running time
        FormSolderPasteMGT.DataGridView1.Columns(6).Width = 180 'end time
        FormSolderPasteMGT.DataGridView1.Columns(7).Width = 120 'status
    End Sub

    Sub KolomCheckRight()
        FormSolderPasteMGT.DataGridView2.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView2.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView2.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView2.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView2.Columns(4).Width = 120 'hole
        FormSolderPasteMGT.DataGridView2.Columns(5).Width = 180 'start time
        FormSolderPasteMGT.DataGridView2.Columns(6).Width = 180 'running time
        FormSolderPasteMGT.DataGridView2.Columns(7).Width = 180 'end time
        FormSolderPasteMGT.DataGridView2.Columns(8).Width = 180 'status
        FormSolderPasteMGT.DataGridView2.Columns(9).Width = 180 'pic
    End Sub

    Sub KolomMixer()
        FormSolderPasteMGT.DataGridView3.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView3.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView3.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView3.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView3.Columns(4).Width = 180 'start time
        FormSolderPasteMGT.DataGridView3.Columns(5).Width = 180 'running time
        FormSolderPasteMGT.DataGridView3.Columns(6).Width = 180 'end time
        FormSolderPasteMGT.DataGridView3.Columns(7).Width = 180 'status
    End Sub

    Sub KolomLine()
        FormSolderPasteMGT.DataGridView4.Columns(0).Width = 150 'Barcode
        FormSolderPasteMGT.DataGridView4.Columns(1).Width = 120 'Maker
        FormSolderPasteMGT.DataGridView4.Columns(2).Width = 120 'Type name
        FormSolderPasteMGT.DataGridView4.Columns(3).Width = 120 'top/bott
        FormSolderPasteMGT.DataGridView4.Columns(4).Width = 120 'line
        FormSolderPasteMGT.DataGridView4.Columns(5).Width = 180 'start time
        FormSolderPasteMGT.DataGridView4.Columns(6).Width = 180 'running time
        FormSolderPasteMGT.DataGridView4.Columns(7).Width = 180 'end time
        FormSolderPasteMGT.DataGridView4.Columns(8).Width = 180 'status
        FormSolderPasteMGT.DataGridView4.Columns(9).Width = 180 'pic
    End Sub

    Sub KolomReadyMixer()
        FormSolderPasteMGT.DataGridView5.Columns(0).Width = 120 'txn id
        FormSolderPasteMGT.DataGridView5.Columns(1).Width = 80 'topbott
        FormSolderPasteMGT.DataGridView5.Columns(2).Width = 60 'hole
    End Sub

    Sub KolomInLine()
        FormSolderPasteMGT.DataGridView6.Columns(0).Width = 120 'txn id
        FormSolderPasteMGT.DataGridView6.Columns(1).Width = 80 'topbott
        FormSolderPasteMGT.DataGridView6.Columns(2).Width = 60 'line
    End Sub

    Sub KolomLogPemakaian()
        FormSolderPasteMGT.DataGridView7.Columns(0).Width = 180 'txn id
        FormSolderPasteMGT.DataGridView7.Columns(1).Width = 180 'topbott
        FormSolderPasteMGT.DataGridView7.Columns(2).Width = 180 'line
        FormSolderPasteMGT.DataGridView7.Columns(3).Width = 180 'start time
        FormSolderPasteMGT.DataGridView7.Columns(4).Width = 180 'run time
        FormSolderPasteMGT.DataGridView7.Columns(5).Width = 180 'end time
        FormSolderPasteMGT.DataGridView7.Columns(6).Width = 180 'pic
    End Sub

    Sub LoadDataCoolingCase()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_coolingcase_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCoolingCase()
    End Sub

    Sub LoadDataCheckRight()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, hole, start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView2.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomCheckRight()
    End Sub

    Sub LoadDataMixer()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, start_time, run_time, end_time, status FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView3.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomMixer()
    End Sub

    Sub LoadDataLine()
        Try
            str = "SELECT txn_id, maker, n_type, topbott, line, start_time, run_time, end_time, status, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView4.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLine()
    End Sub

    Sub LoadDataReadyMixer()
        Try
            str = "SELECT txn_id as Barcode , topbott as Lot, hole as Hole FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView5.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomReadyMixer()
    End Sub

    Sub LoadDataInLine()
        Try
            str = "SELECT txn_id as Barcode , topbott as Lot, line as Line FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'AVAILABLE'"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView6.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomInLine()
    End Sub

    Sub LoadDataLogPemakaian()
        Try
            str = "SELECT txn_id, topbott, line, start_time, run_time, end_time, pic FROM prodsys2_solder_paste_mgt_line_tbl WHERE status = 'EMPTY' ORDER BY end_time DESC"
            Dim da As New MySqlDataAdapter(str, conn)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FormSolderPasteMGT.DataGridView7.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        End Try
        Call KolomLogPemakaian()
    End Sub

    Sub ShowCountCoolingCase()
        FormSolderPasteMGT.Label3.Text = FormSolderPasteMGT.DataGridView1.RowCount
        FormSolderPasteMGT.TabPage1.Text = "Cooling Case [" & FormSolderPasteMGT.DataGridView1.RowCount & "]"
    End Sub

    Sub ShowCountCheckRight()
        FormSolderPasteMGT.Label6.Text = FormSolderPasteMGT.DataGridView2.RowCount
        FormSolderPasteMGT.TabPage2.Text = "Check Right [" & FormSolderPasteMGT.DataGridView2.RowCount & "]"
    End Sub

    Sub ShowCountMixer()
        FormSolderPasteMGT.Label11.Text = FormSolderPasteMGT.DataGridView3.RowCount
        FormSolderPasteMGT.TabPage3.Text = "Mixer [" & FormSolderPasteMGT.DataGridView3.RowCount & "]"
    End Sub

    Sub ShowCountLine()
        FormSolderPasteMGT.Label7.Text = FormSolderPasteMGT.DataGridView4.RowCount
        FormSolderPasteMGT.TabPage4.Text = "Line [" & FormSolderPasteMGT.DataGridView4.RowCount & "]"
    End Sub

    Sub ShowCountReadyMixing()
        FormSolderPasteMGT.Label16.Text = FormSolderPasteMGT.DataGridView5.RowCount
    End Sub

    Sub ShowCountInline()
        FormSolderPasteMGT.Label17.Text = FormSolderPasteMGT.DataGridView6.RowCount
    End Sub

    Sub ShowCountLogPemakaian()
        FormSolderPasteMGT.TabPage5.Text = "History [" & FormSolderPasteMGT.DataGridView7.RowCount & "]"
    End Sub

    Private Sub saveData(sql As String)
        Try
            conn.Open()
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                result = .ExecuteNonQuery()
            End With
            If result > 0 Then
                'MsgBox("Data has been updated in the databse")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub deleteData(sql As String)
        Try
            conn.Open()
            cmd = New MySqlCommand
            With cmd
                .Connection = conn
                .CommandText = sql
                result = .ExecuteNonQuery()
            End With
            If result > 0 Then
                'MsgBox("Data has been deleted in the database")
            End If

        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub FormSolderPasteReturn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            'UPDATE COOLING CASE
            sql = "UPDATE prodsys2_solder_paste_mgt_coolingcase_tbl SET status='AVAILABLE', run_time='', end_time='' WHERE txn_id=" & TextBox1.Text
            saveData(sql)

            ' DELETE CHECKRIGHT
            sql = "DELETE FROM prodsys2_solder_paste_mgt_checkright_tbl WHERE txn_id=" & TextBox1.Text
            deleteData(sql)

            ' DELETE MIXER
            sql = "DELETE FROM prodsys2_solder_paste_mgt_mixer_tbl WHERE txn_id=" & TextBox1.Text
            deleteData(sql)

            ' DELETE LINE
            sql = "DELETE FROM prodsys2_solder_paste_mgt_line_tbl WHERE txn_id=" & TextBox1.Text
            deleteData(sql)

            Call LoadDataCoolingCase()
            Call LoadDataCheckRight()
            Call LoadDataMixer()
            Call LoadDataLine()
            Call LoadDataReadyMixer()
            Call LoadDataInLine()
            Call LoadDataLogPemakaian()

            Call ShowCountCoolingCase()
            Call ShowCountCheckRight()
            Call ShowCountMixer()
            Call ShowCountLine()
            Call ShowCountReadyMixing()
            Call ShowCountInline()
            Call ShowCountLogPemakaian()

            TextBox1.Clear()
            TextBox1.Select()

            MsgBox("SOLDER PASTE BERHASIL DI RETURN KE COOLING CASE!", MsgBoxStyle.Information)
        End If
    End Sub
End Class