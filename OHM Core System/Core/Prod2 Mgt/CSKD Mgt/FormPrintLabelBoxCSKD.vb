Imports Microsoft.Reporting.WinForms

Public Class FormPrintLabelBoxCSKD

    Private Sub FormPrintLabelBoxCSKD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' PRINTED DATE
        Dim datenow As New ReportParameter("datenow", System.DateTime.Now.ToString("dd/MM/yyy HH:mm:ss"))
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {datenow})

        ' TXN ID
        Dim txn_id As New ReportParameter("txn_id", FormLabelBoxCSKD.TextBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {txn_id})

        ' QTY BOX
        Dim qty_box As New ReportParameter("qty_box", FormLabelBoxCSKD.Label22.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {qty_box})

        ' PIC
        Dim pic As New ReportParameter("pic", FormMain.LabelNama.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {pic})

        ' REMARK
        Dim remark As New ReportParameter("remark", FormLabelBoxCSKD.Label15.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {remark})

        ' TYPE BOX
        Dim type_box As New ReportParameter("type_box", FormLabelBoxCSKD.Label13.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {type_box})

        ' PARTNUMBER ASSY
        Dim pn As New ReportParameter("pn", FormLabelBoxCSKD.TextBox3.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {pn})

        ' WEIGHT
        'Dim weight As New ReportParameter("weight", FormLabelBoxCSKD.Label17.Text)
        'ReportViewer1.LocalReport.SetParameters(New ReportParameter() {weight})

        ReportViewer1.ZoomMode = ZoomMode.FullPage
        Me.ReportViewer1.LocalReport.EnableExternalImages = True
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class