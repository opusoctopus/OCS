Imports Microsoft.Reporting.WinForms
Imports MySql.Data.MySqlClient

Public Class FormMSL3PrintLabel
    Private Sub FormMSL3PrintLabel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' TXN ID
        Dim txn_id As New ReportParameter("txn_id", FormMSL3Transaction.TextBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {txn_id})

        ' PARTNUMBER
        Dim pn_lg As New ReportParameter("pn_lg", FormMSL3Transaction.TextBox2.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {pn_lg})

        ' TANGGAL
        Dim tgl As New ReportParameter("tgl", System.DateTime.Now.ToString("dd/MM/yyyy"))
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {tgl})

        ' JAM
        Dim hour As New ReportParameter("hour", System.DateTime.Now.ToString("HH:mm"))
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {hour})

        Me.ReportViewer1.RefreshReport()
    End Sub
End Class