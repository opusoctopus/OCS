<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPrintLabelBoxCSKD
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource3 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.prodsys2_labelbox_cskd_log_tblBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.LabelBoxCSKD = New OHM_Core_System.LabelBoxCSKD()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        CType(Me.prodsys2_labelbox_cskd_log_tblBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LabelBoxCSKD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'prodsys2_labelbox_cskd_log_tblBindingSource
        '
        Me.prodsys2_labelbox_cskd_log_tblBindingSource.DataMember = "prodsys2_labelbox_cskd_log_tbl"
        Me.prodsys2_labelbox_cskd_log_tblBindingSource.DataSource = Me.LabelBoxCSKD
        '
        'LabelBoxCSKD
        '
        Me.LabelBoxCSKD.DataSetName = "LabelBoxCSKD"
        Me.LabelBoxCSKD.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource3.Name = "LabelBoxCSKD"
        ReportDataSource3.Value = Me.prodsys2_labelbox_cskd_log_tblBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource3)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "OHM_Core_System.PrintLabelBoxCSKD.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(678, 805)
        Me.ReportViewer1.TabIndex = 0
        '
        'FormPrintLabelBoxCSKD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(678, 805)
        Me.Controls.Add(Me.ReportViewer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "FormPrintLabelBoxCSKD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormPrintLabelBoxCSKD"
        Me.TopMost = True
        CType(Me.prodsys2_labelbox_cskd_log_tblBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LabelBoxCSKD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents prodsys2_labelbox_cskd_log_tblBindingSource As BindingSource
    Friend WithEvents LabelBoxCSKD As LabelBoxCSKD
End Class
