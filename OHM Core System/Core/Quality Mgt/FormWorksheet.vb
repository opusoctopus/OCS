Imports Microsoft.Reporting.WinForms
Imports MySql.Data.MySqlClient

Public Class FormWorksheet

    Private Sub FormWorksheet_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim title As New ReportParameter("title", "BPR [MANUAL INSERT]")
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {title})

        ' WO
        Dim wo As New ReportParameter("wo", FormBOMKit.TextBox2.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {wo})

        ' ID WORKSHEET
        Dim idworksheet As New ReportParameter("idworksheet", FormBOMKit.TextBox5.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {idworksheet})

        ReportViewer1.ZoomMode = ZoomMode.FullPage
        Me.ReportViewer1.RefreshReport()
    End Sub



End Class