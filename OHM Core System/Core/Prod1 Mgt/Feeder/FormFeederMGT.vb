Imports MySql.Data.MySqlClient

Public Class FormFeederMGT

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
    End Sub

    Sub BuatKolom()
        DataGridView1.Columns(0).Width = 120 'type
        DataGridView1.Columns(1).Width = 120 'sn
        DataGridView1.Columns(2).Width = 120 'Check Point 1
        DataGridView1.Columns(3).Width = 120 'Check Point 2
        DataGridView1.Columns(4).Width = 120 'Check Point 3
        DataGridView1.Columns(5).Width = 120 'Check Point 4
        DataGridView1.Columns(6).Width = 120 'Check Point 5
        DataGridView1.Columns(7).Width = 120 'Check Point 6
        DataGridView1.Columns(8).Width = 120 'Check Point 7
        DataGridView1.Columns(9).Width = 120 'Check Point 8
        DataGridView1.Columns(10).Width = 120 'Check Point 9
        DataGridView1.Columns(11).Width = 120 'Check Point 10
        DataGridView1.Columns(12).Width = 150 'throuble
        DataGridView1.Columns(13).Width = 150 'status
        DataGridView1.Columns(14).Width = 150 'pic
        DataGridView1.Columns(15).Width = 180 'next cal
        DataGridView1.Columns(16).Width = 180 'created
        DataGridView1.Columns(17).Width = 180 'updated
    End Sub

    Sub LoadData()
        Try
            str = "SELECT category as Category, sn as SerialNumber, cp1 as CheckPoint1, cp2 as CheckPoint2, cp3 as CheckPoint3,
cp4 as CheckPoint4, cp5 as CheckPoint5, cp6 as CheckPoint6, cp7 as CheckPoint7, cp8 as CheckPoint8, cp9 as CheckPoint9, cp10 as CheckPoint10, 
throuble as Throuble, status as Status, pic as PIC, next_calibration as NextCalibration, created as Created, updated as Updated FROM prodsys2_feeder_log_tbl"
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

    Sub showcount()
        Label22.Text = DataGridView1.RowCount
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormFeederData.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormFeederInspection.ShowDialog()
    End Sub

    Private Sub FormFeederMGT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
        Call showcount()
    End Sub
End Class