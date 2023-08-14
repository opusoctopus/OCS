Imports MySql.Data.MySqlClient

Public Class FormLocatorFG

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
            .ColumnHeadersHeight = 52
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.White
            .DefaultCellStyle.SelectionForeColor = Color.FromArgb(204, 0, 102)
            With .ColumnHeadersDefaultCellStyle
                .Font = New System.Drawing.Font("Arial", 20.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                .BackColor = Color.FromArgb(135, 156, 167)
                .ForeColor = Color.White
            End With
            With .RowTemplate
                .Height = 40
            End With
        End With
    End Sub

    Sub BuatKolom()
        'DataGridView1.Columns(0).Width = 170 'txnid
        DataGridView1.Columns(0).Width = 220 'locator
        DataGridView1.Columns(1).Width = 210 'wo
        DataGridView1.Columns(2).Width = 250 'pn
        DataGridView1.Columns(3).Width = 450 'model
        DataGridView1.Columns(4).Width = 210 'demand_date
        DataGridView1.Columns(5).Width = 150 'qty
        DataGridView1.Columns(6).Width = 430 'date_IN
        'DataGridView1.Columns(11).Width = 180 'percent
    End Sub

    Sub LoadData()
        Try
            str = "SELECT locator as Locator, wo as WO, pn as PartNumber, model as Model, demand_date as PST, qty as Qty, date_in as DateIn FROM prodsys2_whfg_monitoring_tbl WHERE locator IS NOT NULL ORDER BY wo ASC"
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

    Private Sub FormLocatorFG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call customdgv2()
        Call LoadData()
    End Sub

End Class