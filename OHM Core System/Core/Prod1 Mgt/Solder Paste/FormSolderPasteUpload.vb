Imports MySql.Data.MySqlClient
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient.MySqlDataReader

Public Class FormSolderPasteUpload

    Sub Bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        ListBox1.Items.Clear()
        TextBox4.Text = String.Empty
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Barcode", "Barcode")
        DataGridView1.Columns.Add("Maker", "Maker")
        DataGridView1.Columns.Add("Lot No", "Lot No")
        DataGridView1.Columns.Add("Created", "Created")
        DataGridView1.Columns.Add("Status", "Status")
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).Width = 120
    End Sub

    Private Sub FormSolderPasteUpload_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Bersih()
        Call BuatKolom()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'For i As Integer = Val(TextBox2.Text).ToString To Val(TextBox3.Text).ToString
        'DataGridView1.Rows.Add({TextBox1.Text + (i).ToString, TextBox6.Text, ComboBox1.Text, System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss"), ComboBox2.Text})
        'Next

        Dim OLEcon As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & TextBox4.Text & " ; " & "Extended Properties=Excel 8.0;")
        Dim OLEcmd As New OleDb.OleDbCommand
        Dim OLEda As New OleDb.OleDbDataAdapter
        Dim OLEdt As New DataTable
        Dim sql As String
        Dim resul As Boolean

        Try
            OLEcon.Open()
            With OLEcmd
                .Connection = OLEcon
                .CommandText = "SELECT * FROM [Sheet1$]"
            End With
            OLEda.SelectCommand = OLEcmd
            OLEda.Fill(OLEdt)

            For Each r As DataRow In OLEdt.Rows

                sql = "INSERT INTO prodsys2_solder_paste_mgt_coolingcase_tbl (txn_id,maker,topbott, start_time, status) VALUES ('" & r(0).ToString & "','" & r(1).ToString & "','" & r(2).ToString & "','" & System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss") & "','" & r(3).ToString & "')"
                resul = saveData(sql)
                If resul Then
                    Timer1.Start()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            OLEcon.Close()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Call BuatKolom()

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        ListBox1.Items.Clear()
        TextBox4.Text = String.Empty
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        With OpenFileDialog1
            .Filter = "Excel files(*.xlsx)|*.xlsx|All files (*.*)|*.*"
            .FilterIndex = 1
            .Title = "Import data from Excel file"
        End With
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox4.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Function saveData(sql As String)
        Dim mysqlCOn As MySqlConnection = New MySqlConnection("server=10.217.4.115;user id=ohmuser;password=;database=stockflow_system;sslMode=none")
        Dim mysqlCmd As MySqlCommand
        Dim resul As Boolean

        Try
            mysqlCOn.Open()
            mysqlCmd = New MySqlCommand
            With mysqlCmd
                .Connection = mysqlCOn
                .CommandText = sql
                resul = .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            mysqlCOn.Close()
        End Try
        Return resul
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            MsgBox("Upload Berhasil", MsgBoxStyle.Information, "Sucess!")
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Value += 1
        End If
    End Sub
End Class