Imports Microsoft.Reporting.WinForms
Imports MySql.Data.MySqlClient

Public Class FormWorksheetPreview

    Public Sub SetBindingSource(Datase As DataSet, Table As String, Url As String)
        Dim ReportDataSource3 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim bindingsce As New BindingSource
        bindingsce.DataSource = Datase
        bindingsce.DataMember = "DataTable1"
        ReportDataSource3.Name = "DataSet1"
        ReportDataSource3.Value = bindingsce

        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource
        Dim bindingsce1 As New BindingSource
        bindingsce1.DataSource = Datase
        bindingsce1.DataMember = "DataTable2"
        ReportDataSource1.Name = "DataSet2"
        ReportDataSource1.Value = bindingsce1

        Me.ReportViewer1.LocalReport.DataSources.Clear()
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource3)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = Url
        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub FormWorksheetPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.ReportViewer1.RefreshReport()      

        Dim Datas As New DataSet
        str = "SELECT id_bom, child, description, designators, qty FROM prodsys2_master_bom_tbl WHERE id_bom='" & FormWorksheetPrint.TextBox1.Text & "' AND parent_desc='BPR Insert PCB Assembly' AND supply_type='Assembly Pull'"
        Dim cmd As New MySqlCommand(str, conn)
        da = New MySqlDataAdapter(cmd)
        da.Fill(Datas, "DataTable1")

        str = "SELECT child_sub, description_sub FROM prodsys2_master_bom_sub_tbl WHERE id_bom_sub='" & FormWorksheetPrint.TextBox1.Text & "' AND parent_desc_sub = 'Substitutes'"
        cmd.CommandText = str
        da.SelectCommand = cmd
        da.Fill(Datas, "DataTable2")

        SetBindingSource(Datas, "", "OHM_Core_System.Report1.rdlc")

        ' DATE
        Dim prepare_date As New ReportParameter("prepare_date", System.DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {prepare_date})

        ' LINE
        Dim line As New ReportParameter("line", FormWorksheetPrint.ComboBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {line})

        ' QTY WO
        Dim qtywo As New ReportParameter("qtywo", FormWorksheetPrint.TextBox4.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {qtywo})

        ' ASSY
        Dim assy As New ReportParameter("assy", FormWorksheetPrint.TextBox1.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {assy})

        ' WO
        Dim wo As New ReportParameter("wo", FormWorksheetPrint.TextBox2.Text)
        ReportViewer1.LocalReport.SetParameters(New ReportParameter() {wo})

        Me.ReportViewer1.RefreshReport()
    End Sub
End Class