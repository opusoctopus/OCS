Imports Microsoft.Reporting.WinForms
Public Class FormWorksheetDMS
    Private Sub FormWorksheetDMS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' SHIFT
        If FormBOMKit.CheckBox1.Checked = True Then
            Dim shift As New ReportParameter("shift", FormBOMKit.CheckBox1.Text)
            ReportViewer1.LocalReport.SetParameters(New ReportParameter() {shift})
        ElseIf FormBOMKit.CheckBox2.Checked = True Then
            Dim shift As New ReportParameter("shift", FormBOMKit.CheckBox2.Text)
            ReportViewer1.LocalReport.SetParameters(New ReportParameter() {shift})
        End If

        ' DATE
        Dim datenow As New ReportParameter("datenow", FormMain.LabelDate.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {datenow})

        ' LINE
        Dim line As New ReportParameter("line", FormBOMKit.ComboBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {line})

        ' QTY WO
        Dim qtywo As New ReportParameter("qtywo", FormBOMKit.TextBox4.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {qtywo})

        ' ASSY
        Dim assy As New ReportParameter("assy", FormBOMKit.TextBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {assy})

        ' TITLE
        Dim title As New ReportParameter("title", "DMS [MANUAL INSERT]")
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {title})

        ' WO
        Dim wo As New ReportParameter("wo", FormBOMKit.TextBox3.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {wo})

        ' ID WORKSHEET
        Dim idworksheet As New ReportParameter("idworksheet", FormBOMKit.TextBox5.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {idworksheet})

        ReportViewer1.ZoomMode = ZoomMode.FullPage
        Me.ReportViewer1.RefreshReport()
    End Sub

    Public Sub New(selectedRow2 As List(Of DataGridViewRow))
        InitializeComponent()

        Dim dt As New DataTable("DataTable1")
        dt.Columns.AddRange(New DataColumn(4) {New DataColumn("child", GetType(String)), New DataColumn("description", GetType(String)), New DataColumn("specification", GetType(String)), New DataColumn("designators", GetType(String)), New DataColumn("qty", GetType(String))})
        For Each row As DataGridViewRow In selectedRow2
            dt.Rows.Add(row.Cells("child").Value, row.Cells("description").Value, row.Cells("specification").Value, row.Cells("designators").Value, row.Cells("qty").Value)
        Next

        Dim dsworksheetdms As New Worksheet()
        dsworksheetdms.Merge(dt)
        Dim datasource As New ReportDataSource("WorksheetDMS", dsworksheetdms.Tables(0))
        Me.ReportViewer1.LocalReport.DataSources.Clear()
        Me.ReportViewer1.LocalReport.DataSources.Add(datasource)
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class