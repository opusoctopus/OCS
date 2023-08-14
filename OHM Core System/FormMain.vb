Imports System.IO
Imports System.Reflection

Public Class FormMain
#Region "Jangan Merubah Bagian Ini"
    Public FormActive As String = "FormMain"
    Dim formChild As Form
    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        MsgBox("Developed by Roby Kornela", vbInformation)
    End Sub

    Private Sub btnMenu_Click(sender As Object, e As EventArgs) Handles btnMenu.Click
        PanelSideBar.Visible = Not PanelSideBar.Visible
    End Sub
    Private Sub btnMenu_MouseEnter(sender As Object, e As EventArgs) Handles btnMenu.MouseEnter
        btnMenu.BackColor = Color.FromArgb(200, 5, 83)
    End Sub

    Private Sub btnMenu_MouseLeave(sender As Object, e As EventArgs) Handles btnMenu.MouseLeave
        btnMenu.BackColor = System.Drawing.Color.FromArgb(29, 56, 84)
    End Sub
    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim b = DirectCast(btnLogout, Button)
        b.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 5, 83)
        b = DirectCast(btnMessage, Button)
        b.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 5, 83)
        'Call connection()
        Call loginon()
        TextBox1.Focus()

        LabelDate.Text = FormatDateTime(Now, DateFormat.GeneralDate)

        'Call disabledmenu()
    End Sub

    Public Sub fillFormToPanel(ByVal obj As Form)
        Try
            ProgressBarLoad.Value = 0
            If PanelFormFill.Controls.Find(obj.Name, True).Length = 0 Then
                obj.WindowState = FormWindowState.Normal
                obj.TopLevel = False
                AddHandler obj.Resize, AddressOf Form_Resize
                PanelFormFill.Controls.Add(obj)
                obj.Show()
            Else
                obj = PanelFormFill.Controls.Find(obj.Name, True)(0)
            End If
            If obj.WindowState = FormWindowState.Minimized Then
                obj.WindowState = FormWindowState.Maximized
            End If
            obj.Dock = DockStyle.Fill
            obj.BringToFront()
            LabelTitle.Text = obj.Text
            FormActive = obj.Name
            formChild = obj
            BringFomMinimizeToFront()
            ProgressBarLoad.Value = 100
        Catch ex As Exception
            MsgBox("Form Name Not Found")
        End Try
    End Sub
    Private Sub Form_Resize(sender As Object, e As EventArgs)
        LabelTitle.Text = TryCast(sender, Form).Text
        BringFomMinimizeToFront()
    End Sub
    Sub BringFomMinimizeToFront()
        For Each form In PanelFormFill.Controls
            If form.Tag <> "NoForm" Then
                If form.WindowState = FormWindowState.Minimized Then
                    form.BringToFront()
                End If
            End If
        Next
    End Sub

    Private Sub ButtonMinimize_Click(sender As Object, e As EventArgs) Handles ButtonMinimize.Click
        For Each form In PanelFormFill.Controls
            If form.Tag <> "NoForm" Then
                If form.Visible And form.WindowState <> FormWindowState.Minimized Then
                    form.WindowState = FormWindowState.Minimized
                    form.BringToFront()
                    GoTo end_of_for
                End If
            End If
        Next
end_of_for:
        GetTopFormChild()
    End Sub
    Sub GetTopFormChild()
        For Each form In PanelFormFill.Controls
            If form.Tag <> "NoForm" Then
                If form.Visible And form.WindowState <> FormWindowState.Minimized Then
                    LabelTitle.Text = form.Text
                    FormActive = form.Name
                    GoTo end_of_for
                End If
            End If
        Next
end_of_for:
    End Sub

    Private Sub ButtonClose_Click(sender As Object, e As EventArgs) Handles ButtonClose.Click
        For Each form In PanelFormFill.Controls
            If form.Tag <> "NoForm" Then

                If form.Visible And form.WindowState <> FormWindowState.Minimized Then
                    form.Close()
                    GoTo end_of_for
                End If
            End If
        Next
end_of_for:
        GetTopFormChild()
    End Sub
#End Region
    'Untuk Menampilkan FormChild gunakan perintah fillFormToPanel Sepeti Contoh Berikut

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click
        fillFormToPanel(FormUser)
    End Sub

    Private Sub GroupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GroupToolStripMenuItem.Click
        fillFormToPanel(FormGroup)
    End Sub

    Private Sub AuthToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AuthToolStripMenuItem.Click
        fillFormToPanel(FormRole)
    End Sub

    Private Sub AuthToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AuthToolStripMenuItem1.Click
        fillFormToPanel(FormAuth)
    End Sub

    Private Sub MenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuToolStripMenuItem.Click
        fillFormToPanel(FormScreeningPCBAIDv2)
    End Sub

    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        Call loginon()
        Call bersih()
        LabelSignal.Text = "Not Connected"
        LabelSignal.ForeColor = Color.FromArgb(200, 5, 83)
    End Sub

    Private Sub btnMessage_Click(sender As Object, e As EventArgs) Handles btnMessage.Click
        Me.Close()
    End Sub

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox1.Focus()
    End Sub

    Sub loginoff()
        LabelLogo.Visible = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        Button1.Visible = False
    End Sub

    Sub loginon()
        LabelLogo.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        Button1.Visible = True
        'InventoryToolStripMenuItem1.Enabled = False
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("FORM CAN'T EMPTY!", vbCritical)
                Call bersih()
            Else
                Call connection()
                str = "SELECT * FROM prodsys2_account_tbl WHERE username='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'"
                cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
                rd = cmd.ExecuteReader
                rd.Read()
                Try
                    If rd.HasRows Then
                        Call loginoff()
                        LabelNama.Text = rd("fullname")
                        LabelUserId.Text = rd("dept")
                        LabelSignal.Text = "Connected"
                        LabelSignal.ForeColor = Color.DodgerBlue
                        'Call enabledmenu()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    Else
                        MsgBox("WRONG USERNAME OR PASSWORD!", vbCritical)
                        Call bersih()
                        rd.Close()
                        conn.Close()
                        conn.Dispose()
                    End If
                Catch ex As Exception
                    MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rd.Close()
                    conn.Close()
                    conn.Dispose()
                End Try
            End If
        End If
    End Sub

    Private Sub TimerMain_Tick(sender As Object, e As EventArgs) Handles TimerMain.Tick
        LabelDate.Text = FormatDateTime(Now, DateFormat.GeneralDate)
    End Sub

    Private Sub ScreeningPCBAIDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScreeningPCBAIDToolStripMenuItem.Click
        fillFormToPanel(FormScreeningPCBAID)
    End Sub

    Private Sub MaterialPreparelistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaterialPreparelistToolStripMenuItem.Click
        fillFormToPanel(FormPreparelistMaterial)
    End Sub

    Private Sub WorksheetPrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WorksheetPrintToolStripMenuItem.Click
        fillFormToPanel(FormWorksheetPrint)
    End Sub

    Private Sub ProdPlanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProdPlanToolStripMenuItem.Click
        fillFormToPanel(FormProdPlan)
    End Sub

    Private Sub BPRVerificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BPRVerificationToolStripMenuItem.Click
        fillFormToPanel(FormVerificationBPR)
    End Sub

    Private Sub DMSVerificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DMSVerificationToolStripMenuItem.Click
        fillFormToPanel(FormVerificationDMS)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("FORM CAN'T EMPTY!", vbCritical)
            Call bersih()
        Else
            Call connection()
            str = "SELECT * FROM prodsys2_account_tbl WHERE username='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    Call loginoff()
                    LabelNama.Text = rd("fullname")
                    LabelUserId.Text = rd("dept")
                    LabelSignal.Text = "Connected"
                    LabelSignal.ForeColor = Color.DodgerBlue
                    'Call enabledmenu()
                Else
                    MsgBox("WRONG USERNAME OR PASSWORD!", vbCritical)
                    Call bersih()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR : " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                rd.Close()
                conn.Close()
            End Try
            rd.Close()
            conn.Dispose()
            conn.Close()
        End If
        rd.Close()
        conn.Dispose()
        conn.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles InventoryToolStripMenuItem1.Click
        fillFormToPanel(FormInventory)
    End Sub

    Private Sub ICMappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ICMappingToolStripMenuItem.Click
        fillFormToPanel(FormICMapping)
    End Sub

    Private Sub RequestICEMMCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequestICEMMCToolStripMenuItem.Click
        fillFormToPanel(FormRequestICeMMC)
    End Sub

    Private Sub ApprovalReqICEMMCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApprovalReqICEMMCToolStripMenuItem.Click
        fillFormToPanel(FormApprovalReqICeMMC)
    End Sub

    Private Sub WIMasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WIMasterToolStripMenuItem.Click
        fillFormToPanel(FormMasterWI)
    End Sub

    Private Sub MSL3ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MSL3ToolStripMenuItem.Click
        fillFormToPanel(FormMSL3MGT)
    End Sub

    Private Sub ReturnMaterialNGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReturnMaterialNGToolStripMenuItem.Click
        fillFormToPanel(FormReturnMaterialNG)
    End Sub

    Private Sub LabelBoxCSKDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LabelBoxCSKDToolStripMenuItem.Click
        FormLabelBoxCSKD.Show()
    End Sub

    Sub disabledmenu()
        ToolStripMenuItemSitemAdmin.Enabled = False
        MasterDataToolStripMenuItem.Enabled = False
        MatToolStripMenuItem.Enabled = False
        Prod1ToolStripMenuItem.Enabled = False
        Prod2ToolStripMenuItem.Enabled = False
        QualityToolStripMenuItem.Enabled = False
        DelToolStripMenuItem.Enabled = False
        InventoryToolStripMenuItem1.Enabled = False
        ReportToolStripMenuItem.Enabled = False
    End Sub

    Sub enabledmenu()
        ToolStripMenuItemSitemAdmin.Enabled = True
        MasterDataToolStripMenuItem.Enabled = True
        MatToolStripMenuItem.Enabled = True
        Prod1ToolStripMenuItem.Enabled = True
        Prod2ToolStripMenuItem.Enabled = True
        QualityToolStripMenuItem.Enabled = True
        DelToolStripMenuItem.Enabled = True
        InventoryToolStripMenuItem1.Enabled = True
        ReportToolStripMenuItem.Enabled = True
    End Sub

    Private Sub InspectionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        fillFormToPanel(FormNozzleMGT)
    End Sub

    Private Sub DinoLiteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DinoLiteToolStripMenuItem.Click
        fillFormToPanel(FormDinoLite)
    End Sub

    Private Sub TimbanganCASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimbanganCASToolStripMenuItem.Click
        fillFormToPanel(FormTimbanganCSKD)
    End Sub

    Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Declare Function Wow64EnableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Private osk As String = "C:\Windows\System32\osk.exe"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim old As Long
        If Environment.Is64BitOperatingSystem Then
            If Wow64DisableWow64FsRedirection(old) Then
                Process.Start(osk)
                Wow64EnableWow64FsRedirection(old)
            End If
        Else
            Process.Start(osk)
        End If
    End Sub

    Private Sub DMSInlineVerificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DMSInlineVerificationToolStripMenuItem.Click
        fillFormToPanel(FormVerificationDMSInline)
    End Sub

    Private Sub WIMultiMasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WIMultiMasterToolStripMenuItem.Click
        fillFormToPanel(FormMasterWIMulti)
    End Sub

    Private Sub UploadMasterDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UploadMasterDataToolStripMenuItem.Click
        fillFormToPanel(FormUploadData)
    End Sub

    Private Sub FGMappingLocatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FGMappingLocatorToolStripMenuItem.Click
        fillFormToPanel(FormMappingFG)
    End Sub

    Private Sub LogLabelCSKDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogLabelCSKDToolStripMenuItem.Click
        fillFormToPanel(FormLogLabelBoxCSKD)
    End Sub

    Private Sub RepairCodeMasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RepairCodeMasterToolStripMenuItem.Click
        fillFormToPanel(FormRepairCodeMaster)
    End Sub

    Private Sub RepairPCBAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RepairPCBAToolStripMenuItem.Click
        fillFormToPanel(FormMasterRepair)
    End Sub

    Private Sub FeederManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FeederManagementToolStripMenuItem.Click
        fillFormToPanel(FormFeederMGT)
    End Sub

    Private Sub SolderPasteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SolderPasteToolStripMenuItem1.Click
        FormSolderPasteMGT.Show()
    End Sub

    Private Sub SolderPasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SolderPasteToolStripMenuItem.Click
        FormSolderPasteUpload.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        fillFormToPanel(FormWristtrapMGT)
    End Sub

    Private Sub TransactionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        fillFormToPanel(FormStencilMGT)
    End Sub

    Private Sub InventoryToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormStencilInventory.ShowDialog()
    End Sub

    Private Sub StencilManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StencilManagementToolStripMenuItem.Click
        fillFormToPanel(FormStencilMGT)
    End Sub

    Private Sub NozzleManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NozzleManagementToolStripMenuItem.Click
        fillFormToPanel(FormNozzleMGT)
    End Sub

    Private Sub LabelCompletionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LabelCompletionToolStripMenuItem.Click
        FormCompletion.ShowDialog()
    End Sub

    Private Sub UploadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UploadToolStripMenuItem.Click
        FormMRPUpload.ShowDialog()
    End Sub

    Private Sub MasterDataToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MasterDataToolStripMenuItem1.Click
        fillFormToPanel(FormMRPMaster)
    End Sub

    Private Sub RepairInspectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RepairInspectionToolStripMenuItem.Click
        fillFormToPanel(FormRepairInspection)
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        fillFormToPanel(FormJigMGT)
    End Sub
End Class
