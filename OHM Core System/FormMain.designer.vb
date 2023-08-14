<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    Public Sub New()
        Try
            InitializeComponent()
            MenuApp.Renderer = New MyRender()
        Catch ex As Exception

        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Me.PanelHeader = New System.Windows.Forms.Panel()
        Me.PanelTool = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnMessage = New System.Windows.Forms.Button()
        Me.btnLogout = New System.Windows.Forms.Button()
        Me.LabelTitle = New System.Windows.Forms.Label()
        Me.btnMenu = New System.Windows.Forms.PictureBox()
        Me.PanelFooter = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.PanelMain = New System.Windows.Forms.Panel()
        Me.PanelStatusBar = New System.Windows.Forms.Panel()
        Me.LabelDate = New System.Windows.Forms.Label()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBoxSignal = New System.Windows.Forms.PictureBox()
        Me.LabelSignal = New System.Windows.Forms.Label()
        Me.ProgressBarLoad = New System.Windows.Forms.ProgressBar()
        Me.ButtonMinimize = New System.Windows.Forms.Button()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.PanelFormFill = New System.Windows.Forms.Panel()
        Me.PanelBackGroundImage = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.LabelLogo = New System.Windows.Forms.Label()
        Me.PanelSparator = New System.Windows.Forms.Panel()
        Me.TimerMain = New System.Windows.Forms.Timer(Me.components)
        Me.PanelTitle = New System.Windows.Forms.Panel()
        Me.LabelRefresh = New System.Windows.Forms.Label()
        Me.PanelProfile = New System.Windows.Forms.Panel()
        Me.ButtonConnect = New System.Windows.Forms.Button()
        Me.ButtonHelp = New System.Windows.Forms.Button()
        Me.ButtonSetting = New System.Windows.Forms.Button()
        Me.LabelNama = New System.Windows.Forms.Label()
        Me.ButtonRefresh = New System.Windows.Forms.Button()
        Me.LabelMenu = New System.Windows.Forms.Label()
        Me.PictureBoxProfile = New System.Windows.Forms.PictureBox()
        Me.LabelUserId = New System.Windows.Forms.Label()
        Me.PanelMenu = New System.Windows.Forms.Panel()
        Me.PanelInfo = New System.Windows.Forms.Panel()
        Me.LabelVersion = New System.Windows.Forms.Label()
        Me.MenuApp = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItemSitemAdmin = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AuthToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AuthToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProdASPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProdPlanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SerialNumberToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SerialNumberFALToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WorkInstructionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BOMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadMasterDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MatToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MaterialPreparelistToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MSL3ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StockCardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReturnMaterialNGToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MRPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterDataToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Prod1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BPRVerificationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ICMappingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RequestICEMMCToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StencilManagementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NozzleManagementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SolderPasteManagementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SolderPasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SolderPasteToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FeederManagementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Prod2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DMSVerificationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DMSInlineVerificationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LabelCompletionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LabelBoxCSKDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogLabelCSKDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimbanganCASToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QualityToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ApprovalReqICEMMCToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ScreeningPCBAIDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WorksheetPrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WIMasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WIMultiMasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FGMappingLocatorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InventoryToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RepairPCBAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RepairCodeMasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RepairInspectionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DinoLiteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PanelSideBar = New System.Windows.Forms.Panel()
        Me.PanelHeader.SuspendLayout()
        Me.PanelTool.SuspendLayout()
        CType(Me.btnMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelFooter.SuspendLayout()
        Me.PanelMain.SuspendLayout()
        Me.PanelStatusBar.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxSignal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelFormFill.SuspendLayout()
        Me.PanelBackGroundImage.SuspendLayout()
        Me.PanelTitle.SuspendLayout()
        Me.PanelProfile.SuspendLayout()
        CType(Me.PictureBoxProfile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelMenu.SuspendLayout()
        Me.PanelInfo.SuspendLayout()
        Me.MenuApp.SuspendLayout()
        Me.PanelSideBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelHeader
        '
        Me.PanelHeader.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.PanelHeader.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelHeader.Controls.Add(Me.PanelTool)
        Me.PanelHeader.Controls.Add(Me.LabelTitle)
        Me.PanelHeader.Controls.Add(Me.btnMenu)
        Me.PanelHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelHeader.Location = New System.Drawing.Point(219, 0)
        Me.PanelHeader.Name = "PanelHeader"
        Me.PanelHeader.Size = New System.Drawing.Size(1135, 35)
        Me.PanelHeader.TabIndex = 1
        '
        'PanelTool
        '
        Me.PanelTool.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelTool.BackColor = System.Drawing.Color.Transparent
        Me.PanelTool.Controls.Add(Me.Button2)
        Me.PanelTool.Controls.Add(Me.btnMessage)
        Me.PanelTool.Controls.Add(Me.btnLogout)
        Me.PanelTool.Location = New System.Drawing.Point(952, -3)
        Me.PanelTool.Name = "PanelTool"
        Me.PanelTool.Size = New System.Drawing.Size(183, 42)
        Me.PanelTool.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.BackColor = System.Drawing.Color.Transparent
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button2.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_keyboard_32
        Me.Button2.Location = New System.Drawing.Point(53, -2)
        Me.Button2.Name = "Button2"
        Me.Button2.Padding = New System.Windows.Forms.Padding(3)
        Me.Button2.Size = New System.Drawing.Size(43, 42)
        Me.Button2.TabIndex = 6
        Me.Button2.UseVisualStyleBackColor = False
        '
        'btnMessage
        '
        Me.btnMessage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMessage.BackColor = System.Drawing.Color.Transparent
        Me.btnMessage.FlatAppearance.BorderSize = 0
        Me.btnMessage.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMessage.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnMessage.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_cancel_24
        Me.btnMessage.Location = New System.Drawing.Point(140, -1)
        Me.btnMessage.Name = "btnMessage"
        Me.btnMessage.Padding = New System.Windows.Forms.Padding(3)
        Me.btnMessage.Size = New System.Drawing.Size(43, 41)
        Me.btnMessage.TabIndex = 5
        Me.btnMessage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnMessage.UseVisualStyleBackColor = False
        '
        'btnLogout
        '
        Me.btnLogout.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLogout.BackColor = System.Drawing.Color.Transparent
        Me.btnLogout.FlatAppearance.BorderSize = 0
        Me.btnLogout.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLogout.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogout.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnLogout.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_power_off_button_24
        Me.btnLogout.Location = New System.Drawing.Point(101, -1)
        Me.btnLogout.Name = "btnLogout"
        Me.btnLogout.Padding = New System.Windows.Forms.Padding(3)
        Me.btnLogout.Size = New System.Drawing.Size(40, 41)
        Me.btnLogout.TabIndex = 4
        Me.btnLogout.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnLogout.UseVisualStyleBackColor = False
        '
        'LabelTitle
        '
        Me.LabelTitle.AutoSize = True
        Me.LabelTitle.Font = New System.Drawing.Font("Segoe UI", 12.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTitle.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.LabelTitle.Location = New System.Drawing.Point(43, 6)
        Me.LabelTitle.Name = "LabelTitle"
        Me.LabelTitle.Size = New System.Drawing.Size(157, 23)
        Me.LabelTitle.TabIndex = 3
        Me.LabelTitle.Text = "OHM CORE SYSTEM"
        '
        'btnMenu
        '
        Me.btnMenu.BackColor = System.Drawing.Color.Transparent
        Me.btnMenu.Image = Global.OHM_Core_System.My.Resources.Resources.menu
        Me.btnMenu.Location = New System.Drawing.Point(-1, 0)
        Me.btnMenu.Name = "btnMenu"
        Me.btnMenu.Padding = New System.Windows.Forms.Padding(11)
        Me.btnMenu.Size = New System.Drawing.Size(38, 38)
        Me.btnMenu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.btnMenu.TabIndex = 0
        Me.btnMenu.TabStop = False
        '
        'PanelFooter
        '
        Me.PanelFooter.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.PanelFooter.Controls.Add(Me.Label5)
        Me.PanelFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelFooter.ForeColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.PanelFooter.Location = New System.Drawing.Point(219, 666)
        Me.PanelFooter.Name = "PanelFooter"
        Me.PanelFooter.Size = New System.Drawing.Size(1135, 40)
        Me.PanelFooter.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(705, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(423, 15)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "COPYRIGHT © 2023 | ROBY KORNELA - PT. OHM ELECTRONICS INDONESIA"
        '
        'PanelMain
        '
        Me.PanelMain.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(241, Byte), Integer))
        Me.PanelMain.Controls.Add(Me.PanelStatusBar)
        Me.PanelMain.Controls.Add(Me.ProgressBarLoad)
        Me.PanelMain.Controls.Add(Me.ButtonMinimize)
        Me.PanelMain.Controls.Add(Me.ButtonClose)
        Me.PanelMain.Controls.Add(Me.PanelFormFill)
        Me.PanelMain.Controls.Add(Me.PanelSparator)
        Me.PanelMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelMain.Location = New System.Drawing.Point(219, 35)
        Me.PanelMain.Name = "PanelMain"
        Me.PanelMain.Padding = New System.Windows.Forms.Padding(7, 43, 7, 9)
        Me.PanelMain.Size = New System.Drawing.Size(1135, 631)
        Me.PanelMain.TabIndex = 3
        '
        'PanelStatusBar
        '
        Me.PanelStatusBar.Controls.Add(Me.LabelDate)
        Me.PanelStatusBar.Controls.Add(Me.PictureBox3)
        Me.PanelStatusBar.Controls.Add(Me.PictureBoxSignal)
        Me.PanelStatusBar.Controls.Add(Me.LabelSignal)
        Me.PanelStatusBar.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelStatusBar.Location = New System.Drawing.Point(10, 3)
        Me.PanelStatusBar.Name = "PanelStatusBar"
        Me.PanelStatusBar.Size = New System.Drawing.Size(323, 39)
        Me.PanelStatusBar.TabIndex = 10
        '
        'LabelDate
        '
        Me.LabelDate.AutoSize = True
        Me.LabelDate.ForeColor = System.Drawing.Color.MediumSeaGreen
        Me.LabelDate.Location = New System.Drawing.Point(154, 13)
        Me.LabelDate.Name = "LabelDate"
        Me.LabelDate.Size = New System.Drawing.Size(63, 13)
        Me.LabelDate.TabIndex = 15
        Me.LabelDate.Text = "31/01/2019"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.OHM_Core_System.My.Resources.Resources._date
        Me.PictureBox3.Location = New System.Drawing.Point(129, 10)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(19, 19)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 14
        Me.PictureBox3.TabStop = False
        '
        'PictureBoxSignal
        '
        Me.PictureBoxSignal.Image = Global.OHM_Core_System.My.Resources.Resources.signal_low
        Me.PictureBoxSignal.Location = New System.Drawing.Point(3, 10)
        Me.PictureBoxSignal.Name = "PictureBoxSignal"
        Me.PictureBoxSignal.Size = New System.Drawing.Size(19, 19)
        Me.PictureBoxSignal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxSignal.TabIndex = 8
        Me.PictureBoxSignal.TabStop = False
        '
        'LabelSignal
        '
        Me.LabelSignal.AutoSize = True
        Me.LabelSignal.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSignal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(5, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.LabelSignal.Location = New System.Drawing.Point(28, 13)
        Me.LabelSignal.Name = "LabelSignal"
        Me.LabelSignal.Size = New System.Drawing.Size(85, 13)
        Me.LabelSignal.TabIndex = 9
        Me.LabelSignal.Text = "Not Connected"
        '
        'ProgressBarLoad
        '
        Me.ProgressBarLoad.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBarLoad.Location = New System.Drawing.Point(-13, 0)
        Me.ProgressBarLoad.Name = "ProgressBarLoad"
        Me.ProgressBarLoad.Size = New System.Drawing.Size(1161, 3)
        Me.ProgressBarLoad.TabIndex = 7
        Me.ProgressBarLoad.Value = 100
        '
        'ButtonMinimize
        '
        Me.ButtonMinimize.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonMinimize.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(5, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.ButtonMinimize.FlatAppearance.BorderSize = 0
        Me.ButtonMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonMinimize.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonMinimize.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonMinimize.Image = Global.OHM_Core_System.My.Resources.Resources.minimizeform
        Me.ButtonMinimize.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonMinimize.Location = New System.Drawing.Point(888, 12)
        Me.ButtonMinimize.Name = "ButtonMinimize"
        Me.ButtonMinimize.Size = New System.Drawing.Size(123, 31)
        Me.ButtonMinimize.TabIndex = 6
        Me.ButtonMinimize.Text = "Minimize From"
        Me.ButtonMinimize.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonMinimize.UseVisualStyleBackColor = False
        '
        'ButtonClose
        '
        Me.ButtonClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(5, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.ButtonClose.FlatAppearance.BorderSize = 0
        Me.ButtonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonClose.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonClose.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ButtonClose.Image = Global.OHM_Core_System.My.Resources.Resources.closeform
        Me.ButtonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonClose.Location = New System.Drawing.Point(1019, 12)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(109, 31)
        Me.ButtonClose.TabIndex = 5
        Me.ButtonClose.Text = "Close From"
        Me.ButtonClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonClose.UseVisualStyleBackColor = False
        '
        'PanelFormFill
        '
        Me.PanelFormFill.BackColor = System.Drawing.Color.White
        Me.PanelFormFill.Controls.Add(Me.PanelBackGroundImage)
        Me.PanelFormFill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelFormFill.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelFormFill.Location = New System.Drawing.Point(7, 52)
        Me.PanelFormFill.Margin = New System.Windows.Forms.Padding(17)
        Me.PanelFormFill.Name = "PanelFormFill"
        Me.PanelFormFill.Size = New System.Drawing.Size(1121, 570)
        Me.PanelFormFill.TabIndex = 4
        '
        'PanelBackGroundImage
        '
        Me.PanelBackGroundImage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelBackGroundImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelBackGroundImage.Controls.Add(Me.Button1)
        Me.PanelBackGroundImage.Controls.Add(Me.TextBox2)
        Me.PanelBackGroundImage.Controls.Add(Me.TextBox1)
        Me.PanelBackGroundImage.Controls.Add(Me.LabelLogo)
        Me.PanelBackGroundImage.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelBackGroundImage.Location = New System.Drawing.Point(3, 6)
        Me.PanelBackGroundImage.Name = "PanelBackGroundImage"
        Me.PanelBackGroundImage.Size = New System.Drawing.Size(1115, 561)
        Me.PanelBackGroundImage.TabIndex = 8
        Me.PanelBackGroundImage.Tag = "NoForm"
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(135, Byte), Integer), CType(CType(156, Byte), Integer), CType(CType(167, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(371, 360)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(204, 32)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "L O G I N"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Location = New System.Drawing.Point(371, 311)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9787)
        Me.TextBox2.Size = New System.Drawing.Size(204, 27)
        Me.TextBox2.TabIndex = 10
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(371, 268)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(204, 27)
        Me.TextBox1.TabIndex = 9
        '
        'LabelLogo
        '
        Me.LabelLogo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelLogo.BackColor = System.Drawing.Color.Transparent
        Me.LabelLogo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LabelLogo.Font = New System.Drawing.Font("Segoe UI", 18.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelLogo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(5, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.LabelLogo.Location = New System.Drawing.Point(371, 219)
        Me.LabelLogo.Name = "LabelLogo"
        Me.LabelLogo.Size = New System.Drawing.Size(204, 35)
        Me.LabelLogo.TabIndex = 8
        Me.LabelLogo.Tag = "NoForm"
        Me.LabelLogo.Text = "Login System"
        Me.LabelLogo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PanelSparator
        '
        Me.PanelSparator.BackColor = System.Drawing.Color.FromArgb(CType(CType(200, Byte), Integer), CType(CType(5, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.PanelSparator.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelSparator.Location = New System.Drawing.Point(7, 43)
        Me.PanelSparator.Name = "PanelSparator"
        Me.PanelSparator.Size = New System.Drawing.Size(1121, 9)
        Me.PanelSparator.TabIndex = 0
        '
        'TimerMain
        '
        Me.TimerMain.Enabled = True
        Me.TimerMain.Interval = 1000
        '
        'PanelTitle
        '
        Me.PanelTitle.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.PanelTitle.Controls.Add(Me.LabelRefresh)
        Me.PanelTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelTitle.Location = New System.Drawing.Point(0, 0)
        Me.PanelTitle.Name = "PanelTitle"
        Me.PanelTitle.Size = New System.Drawing.Size(219, 38)
        Me.PanelTitle.TabIndex = 1
        '
        'LabelRefresh
        '
        Me.LabelRefresh.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelRefresh.AutoSize = True
        Me.LabelRefresh.Font = New System.Drawing.Font("Segoe UI", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelRefresh.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.LabelRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelRefresh.Location = New System.Drawing.Point(48, 3)
        Me.LabelRefresh.Name = "LabelRefresh"
        Me.LabelRefresh.Size = New System.Drawing.Size(124, 28)
        Me.LabelRefresh.TabIndex = 0
        Me.LabelRefresh.Text = "App System"
        Me.LabelRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PanelProfile
        '
        Me.PanelProfile.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.PanelProfile.Controls.Add(Me.ButtonConnect)
        Me.PanelProfile.Controls.Add(Me.ButtonHelp)
        Me.PanelProfile.Controls.Add(Me.ButtonSetting)
        Me.PanelProfile.Controls.Add(Me.LabelNama)
        Me.PanelProfile.Controls.Add(Me.ButtonRefresh)
        Me.PanelProfile.Controls.Add(Me.LabelMenu)
        Me.PanelProfile.Controls.Add(Me.PictureBoxProfile)
        Me.PanelProfile.Controls.Add(Me.LabelUserId)
        Me.PanelProfile.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelProfile.Location = New System.Drawing.Point(0, 38)
        Me.PanelProfile.Name = "PanelProfile"
        Me.PanelProfile.Size = New System.Drawing.Size(219, 162)
        Me.PanelProfile.TabIndex = 5
        '
        'ButtonConnect
        '
        Me.ButtonConnect.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonConnect.FlatAppearance.BorderSize = 0
        Me.ButtonConnect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonConnect.Image = Global.OHM_Core_System.My.Resources.Resources.connect
        Me.ButtonConnect.Location = New System.Drawing.Point(59, 128)
        Me.ButtonConnect.Name = "ButtonConnect"
        Me.ButtonConnect.Size = New System.Drawing.Size(38, 31)
        Me.ButtonConnect.TabIndex = 6
        Me.ButtonConnect.UseVisualStyleBackColor = True
        '
        'ButtonHelp
        '
        Me.ButtonHelp.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonHelp.FlatAppearance.BorderSize = 0
        Me.ButtonHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonHelp.Image = Global.OHM_Core_System.My.Resources.Resources.help
        Me.ButtonHelp.Location = New System.Drawing.Point(6, 128)
        Me.ButtonHelp.Name = "ButtonHelp"
        Me.ButtonHelp.Size = New System.Drawing.Size(38, 31)
        Me.ButtonHelp.TabIndex = 5
        Me.ButtonHelp.UseVisualStyleBackColor = True
        '
        'ButtonSetting
        '
        Me.ButtonSetting.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonSetting.FlatAppearance.BorderSize = 0
        Me.ButtonSetting.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonSetting.Image = Global.OHM_Core_System.My.Resources.Resources.setting
        Me.ButtonSetting.Location = New System.Drawing.Point(175, 128)
        Me.ButtonSetting.Name = "ButtonSetting"
        Me.ButtonSetting.Padding = New System.Windows.Forms.Padding(3)
        Me.ButtonSetting.Size = New System.Drawing.Size(38, 31)
        Me.ButtonSetting.TabIndex = 2
        Me.ButtonSetting.UseVisualStyleBackColor = True
        '
        'LabelNama
        '
        Me.LabelNama.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelNama.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNama.ForeColor = System.Drawing.Color.White
        Me.LabelNama.Location = New System.Drawing.Point(3, 68)
        Me.LabelNama.Name = "LabelNama"
        Me.LabelNama.Size = New System.Drawing.Size(213, 16)
        Me.LabelNama.TabIndex = 4
        Me.LabelNama.Text = "Anonymous"
        Me.LabelNama.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ButtonRefresh.FlatAppearance.BorderSize = 0
        Me.ButtonRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonRefresh.Image = Global.OHM_Core_System.My.Resources.Resources.refresh
        Me.ButtonRefresh.Location = New System.Drawing.Point(116, 128)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(38, 31)
        Me.ButtonRefresh.TabIndex = 1
        Me.ButtonRefresh.UseVisualStyleBackColor = True
        '
        'LabelMenu
        '
        Me.LabelMenu.AutoSize = True
        Me.LabelMenu.Font = New System.Drawing.Font("Microsoft YaHei UI", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMenu.ForeColor = System.Drawing.Color.White
        Me.LabelMenu.Location = New System.Drawing.Point(56, 103)
        Me.LabelMenu.Name = "LabelMenu"
        Me.LabelMenu.Size = New System.Drawing.Size(99, 24)
        Me.LabelMenu.TabIndex = 1
        Me.LabelMenu.Text = "Menu App"
        Me.LabelMenu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBoxProfile
        '
        Me.PictureBoxProfile.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.PictureBoxProfile.Image = Global.OHM_Core_System.My.Resources.Resources.user
        Me.PictureBoxProfile.Location = New System.Drawing.Point(74, 6)
        Me.PictureBoxProfile.Name = "PictureBoxProfile"
        Me.PictureBoxProfile.Size = New System.Drawing.Size(60, 59)
        Me.PictureBoxProfile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxProfile.TabIndex = 2
        Me.PictureBoxProfile.TabStop = False
        '
        'LabelUserId
        '
        Me.LabelUserId.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelUserId.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUserId.ForeColor = System.Drawing.Color.White
        Me.LabelUserId.Location = New System.Drawing.Point(3, 87)
        Me.LabelUserId.Name = "LabelUserId"
        Me.LabelUserId.Size = New System.Drawing.Size(213, 16)
        Me.LabelUserId.TabIndex = 3
        Me.LabelUserId.Text = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        Me.LabelUserId.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PanelMenu
        '
        Me.PanelMenu.AutoScroll = True
        Me.PanelMenu.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.PanelMenu.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelMenu.Controls.Add(Me.PanelInfo)
        Me.PanelMenu.Controls.Add(Me.MenuApp)
        Me.PanelMenu.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelMenu.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelMenu.Location = New System.Drawing.Point(0, 200)
        Me.PanelMenu.Margin = New System.Windows.Forms.Padding(9, 3, 9, 3)
        Me.PanelMenu.Name = "PanelMenu"
        Me.PanelMenu.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.PanelMenu.Size = New System.Drawing.Size(219, 506)
        Me.PanelMenu.TabIndex = 5
        '
        'PanelInfo
        '
        Me.PanelInfo.Controls.Add(Me.LabelVersion)
        Me.PanelInfo.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelInfo.Location = New System.Drawing.Point(0, 466)
        Me.PanelInfo.Name = "PanelInfo"
        Me.PanelInfo.Size = New System.Drawing.Size(219, 40)
        Me.PanelInfo.TabIndex = 1
        '
        'LabelVersion
        '
        Me.LabelVersion.BackColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(84, Byte), Integer))
        Me.LabelVersion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelVersion.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVersion.ForeColor = System.Drawing.Color.White
        Me.LabelVersion.Location = New System.Drawing.Point(0, 0)
        Me.LabelVersion.Name = "LabelVersion"
        Me.LabelVersion.Size = New System.Drawing.Size(219, 40)
        Me.LabelVersion.TabIndex = 1
        Me.LabelVersion.Text = "V2.6 Build 0809"
        Me.LabelVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MenuApp
        '
        Me.MenuApp.BackColor = System.Drawing.Color.Transparent
        Me.MenuApp.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuApp.GripMargin = New System.Windows.Forms.Padding(19)
        Me.MenuApp.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemSitemAdmin, Me.MasterDataToolStripMenuItem, Me.MatToolStripMenuItem, Me.Prod1ToolStripMenuItem, Me.Prod2ToolStripMenuItem, Me.QualityToolStripMenuItem, Me.DelToolStripMenuItem, Me.InventoryToolStripMenuItem1, Me.ReportToolStripMenuItem, Me.ToolStripMenuItem3, Me.ToolsToolStripMenuItem, Me.ToolStripMenuItem1})
        Me.MenuApp.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.VerticalStackWithOverflow
        Me.MenuApp.Location = New System.Drawing.Point(0, 1)
        Me.MenuApp.Name = "MenuApp"
        Me.MenuApp.Padding = New System.Windows.Forms.Padding(1)
        Me.MenuApp.Size = New System.Drawing.Size(219, 292)
        Me.MenuApp.TabIndex = 0
        Me.MenuApp.Text = "MenuStrip"
        '
        'ToolStripMenuItemSitemAdmin
        '
        Me.ToolStripMenuItemSitemAdmin.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserToolStripMenuItem, Me.GroupToolStripMenuItem, Me.AuthToolStripMenuItem, Me.AuthToolStripMenuItem1, Me.MenuToolStripMenuItem})
        Me.ToolStripMenuItemSitemAdmin.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_circled_menu_24
        Me.ToolStripMenuItemSitemAdmin.Name = "ToolStripMenuItemSitemAdmin"
        Me.ToolStripMenuItemSitemAdmin.Size = New System.Drawing.Size(216, 24)
        Me.ToolStripMenuItemSitemAdmin.Text = "Sistem Admin"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_menu_vertical_24
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(123, 24)
        Me.UserToolStripMenuItem.Text = "User"
        '
        'GroupToolStripMenuItem
        '
        Me.GroupToolStripMenuItem.Name = "GroupToolStripMenuItem"
        Me.GroupToolStripMenuItem.Size = New System.Drawing.Size(123, 24)
        Me.GroupToolStripMenuItem.Text = "Group"
        '
        'AuthToolStripMenuItem
        '
        Me.AuthToolStripMenuItem.Name = "AuthToolStripMenuItem"
        Me.AuthToolStripMenuItem.Size = New System.Drawing.Size(123, 24)
        Me.AuthToolStripMenuItem.Text = "Role"
        '
        'AuthToolStripMenuItem1
        '
        Me.AuthToolStripMenuItem1.Name = "AuthToolStripMenuItem1"
        Me.AuthToolStripMenuItem1.Size = New System.Drawing.Size(123, 24)
        Me.AuthToolStripMenuItem1.Text = "Auth"
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        Me.MenuToolStripMenuItem.Size = New System.Drawing.Size(123, 24)
        Me.MenuToolStripMenuItem.Text = "Menu"
        '
        'MasterDataToolStripMenuItem
        '
        Me.MasterDataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProdASPToolStripMenuItem, Me.ProdPlanToolStripMenuItem, Me.SerialNumberToolStripMenuItem, Me.SerialNumberFALToolStripMenuItem, Me.WorkInstructionToolStripMenuItem, Me.BOMToolStripMenuItem, Me.UploadMasterDataToolStripMenuItem})
        Me.MasterDataToolStripMenuItem.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_circled_menu_24
        Me.MasterDataToolStripMenuItem.Name = "MasterDataToolStripMenuItem"
        Me.MasterDataToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.MasterDataToolStripMenuItem.Text = "Master Data"
        '
        'ProdASPToolStripMenuItem
        '
        Me.ProdASPToolStripMenuItem.Name = "ProdASPToolStripMenuItem"
        Me.ProdASPToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.ProdASPToolStripMenuItem.Text = "Prod. ASP"
        '
        'ProdPlanToolStripMenuItem
        '
        Me.ProdPlanToolStripMenuItem.Name = "ProdPlanToolStripMenuItem"
        Me.ProdPlanToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.ProdPlanToolStripMenuItem.Text = "Prod. Plan"
        '
        'SerialNumberToolStripMenuItem
        '
        Me.SerialNumberToolStripMenuItem.Name = "SerialNumberToolStripMenuItem"
        Me.SerialNumberToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.SerialNumberToolStripMenuItem.Text = "Serial Number"
        '
        'SerialNumberFALToolStripMenuItem
        '
        Me.SerialNumberFALToolStripMenuItem.Name = "SerialNumberFALToolStripMenuItem"
        Me.SerialNumberFALToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.SerialNumberFALToolStripMenuItem.Text = "Serial Number FAL"
        '
        'WorkInstructionToolStripMenuItem
        '
        Me.WorkInstructionToolStripMenuItem.Name = "WorkInstructionToolStripMenuItem"
        Me.WorkInstructionToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.WorkInstructionToolStripMenuItem.Text = "Works Instruction"
        '
        'BOMToolStripMenuItem
        '
        Me.BOMToolStripMenuItem.Name = "BOMToolStripMenuItem"
        Me.BOMToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.BOMToolStripMenuItem.Text = "BOM"
        '
        'UploadMasterDataToolStripMenuItem
        '
        Me.UploadMasterDataToolStripMenuItem.Name = "UploadMasterDataToolStripMenuItem"
        Me.UploadMasterDataToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.UploadMasterDataToolStripMenuItem.Text = "Upload Master Data"
        '
        'MatToolStripMenuItem
        '
        Me.MatToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MaterialPreparelistToolStripMenuItem, Me.MSL3ToolStripMenuItem, Me.StockCardToolStripMenuItem, Me.ReturnMaterialNGToolStripMenuItem, Me.MRPToolStripMenuItem})
        Me.MatToolStripMenuItem.Image = Global.OHM_Core_System.My.Resources.Resources.icons8_circled_menu_24
        Me.MatToolStripMenuItem.Name = "MatToolStripMenuItem"
        Me.MatToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.MatToolStripMenuItem.Text = "Material Management"
        '
        'MaterialPreparelistToolStripMenuItem
        '
        Me.MaterialPreparelistToolStripMenuItem.Name = "MaterialPreparelistToolStripMenuItem"
        Me.MaterialPreparelistToolStripMenuItem.Size = New System.Drawing.Size(225, 24)
        Me.MaterialPreparelistToolStripMenuItem.Text = "Material Preparelist"
        '
        'MSL3ToolStripMenuItem
        '
        Me.MSL3ToolStripMenuItem.Name = "MSL3ToolStripMenuItem"
        Me.MSL3ToolStripMenuItem.Size = New System.Drawing.Size(225, 24)
        Me.MSL3ToolStripMenuItem.Text = "MSL 3"
        '
        'StockCardToolStripMenuItem
        '
        Me.StockCardToolStripMenuItem.Name = "StockCardToolStripMenuItem"
        Me.StockCardToolStripMenuItem.Size = New System.Drawing.Size(225, 24)
        Me.StockCardToolStripMenuItem.Text = "Stock Card"
        '
        'ReturnMaterialNGToolStripMenuItem
        '
        Me.ReturnMaterialNGToolStripMenuItem.Name = "ReturnMaterialNGToolStripMenuItem"
        Me.ReturnMaterialNGToolStripMenuItem.Size = New System.Drawing.Size(225, 24)
        Me.ReturnMaterialNGToolStripMenuItem.Text = "Return Material (NG)"
        '
        'MRPToolStripMenuItem
        '
        Me.MRPToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MasterDataToolStripMenuItem1, Me.UploadToolStripMenuItem})
        Me.MRPToolStripMenuItem.Name = "MRPToolStripMenuItem"
        Me.MRPToolStripMenuItem.Size = New System.Drawing.Size(225, 24)
        Me.MRPToolStripMenuItem.Text = "MRP"
        '
        'MasterDataToolStripMenuItem1
        '
        Me.MasterDataToolStripMenuItem1.Name = "MasterDataToolStripMenuItem1"
        Me.MasterDataToolStripMenuItem1.Size = New System.Drawing.Size(166, 24)
        Me.MasterDataToolStripMenuItem1.Text = "Master Data"
        '
        'UploadToolStripMenuItem
        '
        Me.UploadToolStripMenuItem.Name = "UploadToolStripMenuItem"
        Me.UploadToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.UploadToolStripMenuItem.Text = "Upload"
        '
        'Prod1ToolStripMenuItem
        '
        Me.Prod1ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BPRVerificationToolStripMenuItem, Me.ICMappingToolStripMenuItem, Me.RequestICEMMCToolStripMenuItem, Me.StencilManagementToolStripMenuItem, Me.NozzleManagementToolStripMenuItem, Me.SolderPasteManagementToolStripMenuItem, Me.FeederManagementToolStripMenuItem})
        Me.Prod1ToolStripMenuItem.Image = CType(resources.GetObject("Prod1ToolStripMenuItem.Image"), System.Drawing.Image)
        Me.Prod1ToolStripMenuItem.Name = "Prod1ToolStripMenuItem"
        Me.Prod1ToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.Prod1ToolStripMenuItem.Text = "Prod. 1 Management"
        '
        'BPRVerificationToolStripMenuItem
        '
        Me.BPRVerificationToolStripMenuItem.Name = "BPRVerificationToolStripMenuItem"
        Me.BPRVerificationToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.BPRVerificationToolStripMenuItem.Text = "BPR Verification"
        '
        'ICMappingToolStripMenuItem
        '
        Me.ICMappingToolStripMenuItem.Name = "ICMappingToolStripMenuItem"
        Me.ICMappingToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.ICMappingToolStripMenuItem.Text = "IC Mapping"
        '
        'RequestICEMMCToolStripMenuItem
        '
        Me.RequestICEMMCToolStripMenuItem.Name = "RequestICEMMCToolStripMenuItem"
        Me.RequestICEMMCToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.RequestICEMMCToolStripMenuItem.Text = "Request IC eMMC"
        '
        'StencilManagementToolStripMenuItem
        '
        Me.StencilManagementToolStripMenuItem.Name = "StencilManagementToolStripMenuItem"
        Me.StencilManagementToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.StencilManagementToolStripMenuItem.Text = "Stencil Management"
        '
        'NozzleManagementToolStripMenuItem
        '
        Me.NozzleManagementToolStripMenuItem.Name = "NozzleManagementToolStripMenuItem"
        Me.NozzleManagementToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.NozzleManagementToolStripMenuItem.Text = "Nozzle Management"
        '
        'SolderPasteManagementToolStripMenuItem
        '
        Me.SolderPasteManagementToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SolderPasteToolStripMenuItem, Me.SolderPasteToolStripMenuItem1})
        Me.SolderPasteManagementToolStripMenuItem.Name = "SolderPasteManagementToolStripMenuItem"
        Me.SolderPasteManagementToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.SolderPasteManagementToolStripMenuItem.Text = "Solder Paste Management"
        '
        'SolderPasteToolStripMenuItem
        '
        Me.SolderPasteToolStripMenuItem.Name = "SolderPasteToolStripMenuItem"
        Me.SolderPasteToolStripMenuItem.Size = New System.Drawing.Size(169, 24)
        Me.SolderPasteToolStripMenuItem.Text = "Upload Data"
        '
        'SolderPasteToolStripMenuItem1
        '
        Me.SolderPasteToolStripMenuItem1.Name = "SolderPasteToolStripMenuItem1"
        Me.SolderPasteToolStripMenuItem1.Size = New System.Drawing.Size(169, 24)
        Me.SolderPasteToolStripMenuItem1.Text = "Solder Paste"
        '
        'FeederManagementToolStripMenuItem
        '
        Me.FeederManagementToolStripMenuItem.Name = "FeederManagementToolStripMenuItem"
        Me.FeederManagementToolStripMenuItem.Size = New System.Drawing.Size(267, 24)
        Me.FeederManagementToolStripMenuItem.Text = "Feeder Management"
        '
        'Prod2ToolStripMenuItem
        '
        Me.Prod2ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DMSVerificationToolStripMenuItem, Me.DMSInlineVerificationToolStripMenuItem, Me.LabelCompletionToolStripMenuItem, Me.LabelBoxCSKDToolStripMenuItem, Me.LogLabelCSKDToolStripMenuItem, Me.TimbanganCASToolStripMenuItem})
        Me.Prod2ToolStripMenuItem.Image = CType(resources.GetObject("Prod2ToolStripMenuItem.Image"), System.Drawing.Image)
        Me.Prod2ToolStripMenuItem.Name = "Prod2ToolStripMenuItem"
        Me.Prod2ToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.Prod2ToolStripMenuItem.Text = "Prod. 2 Management"
        '
        'DMSVerificationToolStripMenuItem
        '
        Me.DMSVerificationToolStripMenuItem.Name = "DMSVerificationToolStripMenuItem"
        Me.DMSVerificationToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.DMSVerificationToolStripMenuItem.Text = "DMS Verification"
        '
        'DMSInlineVerificationToolStripMenuItem
        '
        Me.DMSInlineVerificationToolStripMenuItem.Name = "DMSInlineVerificationToolStripMenuItem"
        Me.DMSInlineVerificationToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.DMSInlineVerificationToolStripMenuItem.Text = "DMS Inline Verification"
        '
        'LabelCompletionToolStripMenuItem
        '
        Me.LabelCompletionToolStripMenuItem.Name = "LabelCompletionToolStripMenuItem"
        Me.LabelCompletionToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.LabelCompletionToolStripMenuItem.Text = "Label Completion"
        '
        'LabelBoxCSKDToolStripMenuItem
        '
        Me.LabelBoxCSKDToolStripMenuItem.Name = "LabelBoxCSKDToolStripMenuItem"
        Me.LabelBoxCSKDToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.LabelBoxCSKDToolStripMenuItem.Text = "Label Box CSKD"
        '
        'LogLabelCSKDToolStripMenuItem
        '
        Me.LogLabelCSKDToolStripMenuItem.Name = "LogLabelCSKDToolStripMenuItem"
        Me.LogLabelCSKDToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.LogLabelCSKDToolStripMenuItem.Text = "Log Label CSKD"
        '
        'TimbanganCASToolStripMenuItem
        '
        Me.TimbanganCASToolStripMenuItem.Name = "TimbanganCASToolStripMenuItem"
        Me.TimbanganCASToolStripMenuItem.Size = New System.Drawing.Size(239, 24)
        Me.TimbanganCASToolStripMenuItem.Text = "Timbangan CAS"
        '
        'QualityToolStripMenuItem
        '
        Me.QualityToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApprovalReqICEMMCToolStripMenuItem, Me.ScreeningPCBAIDToolStripMenuItem, Me.WorksheetPrintToolStripMenuItem, Me.WIMasterToolStripMenuItem, Me.WIMultiMasterToolStripMenuItem})
        Me.QualityToolStripMenuItem.Image = CType(resources.GetObject("QualityToolStripMenuItem.Image"), System.Drawing.Image)
        Me.QualityToolStripMenuItem.Name = "QualityToolStripMenuItem"
        Me.QualityToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.QualityToolStripMenuItem.Text = "Quality Management"
        '
        'ApprovalReqICEMMCToolStripMenuItem
        '
        Me.ApprovalReqICEMMCToolStripMenuItem.Name = "ApprovalReqICEMMCToolStripMenuItem"
        Me.ApprovalReqICEMMCToolStripMenuItem.Size = New System.Drawing.Size(248, 24)
        Me.ApprovalReqICEMMCToolStripMenuItem.Text = "Approval Req. IC eMMC"
        '
        'ScreeningPCBAIDToolStripMenuItem
        '
        Me.ScreeningPCBAIDToolStripMenuItem.Name = "ScreeningPCBAIDToolStripMenuItem"
        Me.ScreeningPCBAIDToolStripMenuItem.Size = New System.Drawing.Size(248, 24)
        Me.ScreeningPCBAIDToolStripMenuItem.Text = "Screening PCBA ID"
        '
        'WorksheetPrintToolStripMenuItem
        '
        Me.WorksheetPrintToolStripMenuItem.Name = "WorksheetPrintToolStripMenuItem"
        Me.WorksheetPrintToolStripMenuItem.Size = New System.Drawing.Size(248, 24)
        Me.WorksheetPrintToolStripMenuItem.Text = "Worksheet Print"
        '
        'WIMasterToolStripMenuItem
        '
        Me.WIMasterToolStripMenuItem.Name = "WIMasterToolStripMenuItem"
        Me.WIMasterToolStripMenuItem.Size = New System.Drawing.Size(248, 24)
        Me.WIMasterToolStripMenuItem.Text = "WI Master"
        '
        'WIMultiMasterToolStripMenuItem
        '
        Me.WIMultiMasterToolStripMenuItem.Name = "WIMultiMasterToolStripMenuItem"
        Me.WIMultiMasterToolStripMenuItem.Size = New System.Drawing.Size(248, 24)
        Me.WIMultiMasterToolStripMenuItem.Text = "WI Multi Master"
        '
        'DelToolStripMenuItem
        '
        Me.DelToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FGMappingLocatorToolStripMenuItem})
        Me.DelToolStripMenuItem.Image = CType(resources.GetObject("DelToolStripMenuItem.Image"), System.Drawing.Image)
        Me.DelToolStripMenuItem.Name = "DelToolStripMenuItem"
        Me.DelToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.DelToolStripMenuItem.Text = "Delivery Management"
        '
        'FGMappingLocatorToolStripMenuItem
        '
        Me.FGMappingLocatorToolStripMenuItem.Name = "FGMappingLocatorToolStripMenuItem"
        Me.FGMappingLocatorToolStripMenuItem.Size = New System.Drawing.Size(228, 24)
        Me.FGMappingLocatorToolStripMenuItem.Text = "F/G Mapping Locator"
        '
        'InventoryToolStripMenuItem1
        '
        Me.InventoryToolStripMenuItem1.Image = CType(resources.GetObject("InventoryToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.InventoryToolStripMenuItem1.Name = "InventoryToolStripMenuItem1"
        Me.InventoryToolStripMenuItem1.Size = New System.Drawing.Size(216, 24)
        Me.InventoryToolStripMenuItem1.Text = "Inventory Assy Mgt"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RepairPCBAToolStripMenuItem, Me.RepairCodeMasterToolStripMenuItem, Me.RepairInspectionToolStripMenuItem})
        Me.ReportToolStripMenuItem.Image = CType(resources.GetObject("ReportToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.ReportToolStripMenuItem.Text = "Repair Management"
        '
        'RepairPCBAToolStripMenuItem
        '
        Me.RepairPCBAToolStripMenuItem.Name = "RepairPCBAToolStripMenuItem"
        Me.RepairPCBAToolStripMenuItem.Size = New System.Drawing.Size(220, 24)
        Me.RepairPCBAToolStripMenuItem.Text = "Repair PCBA"
        '
        'RepairCodeMasterToolStripMenuItem
        '
        Me.RepairCodeMasterToolStripMenuItem.Name = "RepairCodeMasterToolStripMenuItem"
        Me.RepairCodeMasterToolStripMenuItem.Size = New System.Drawing.Size(220, 24)
        Me.RepairCodeMasterToolStripMenuItem.Text = "Repair Code Master"
        '
        'RepairInspectionToolStripMenuItem
        '
        Me.RepairInspectionToolStripMenuItem.Name = "RepairInspectionToolStripMenuItem"
        Me.RepairInspectionToolStripMenuItem.Size = New System.Drawing.Size(220, 24)
        Me.RepairInspectionToolStripMenuItem.Text = "Repair Inspection"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem2})
        Me.ToolStripMenuItem3.Image = CType(resources.GetObject("ToolStripMenuItem3.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(216, 24)
        Me.ToolStripMenuItem3.Text = "Engineering Management"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(139, 24)
        Me.ToolStripMenuItem2.Text = "DFT JIG"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DinoLiteToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Image = CType(resources.GetObject("ToolsToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(216, 24)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'DinoLiteToolStripMenuItem
        '
        Me.DinoLiteToolStripMenuItem.Name = "DinoLiteToolStripMenuItem"
        Me.DinoLiteToolStripMenuItem.Size = New System.Drawing.Size(141, 24)
        Me.DinoLiteToolStripMenuItem.Text = "Dino Lite"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Image = CType(resources.GetObject("ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(216, 24)
        Me.ToolStripMenuItem1.Text = "Wrist Strap Checker"
        '
        'PanelSideBar
        '
        Me.PanelSideBar.BackColor = System.Drawing.Color.MediumTurquoise
        Me.PanelSideBar.Controls.Add(Me.PanelMenu)
        Me.PanelSideBar.Controls.Add(Me.PanelProfile)
        Me.PanelSideBar.Controls.Add(Me.PanelTitle)
        Me.PanelSideBar.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelSideBar.Location = New System.Drawing.Point(0, 0)
        Me.PanelSideBar.Name = "PanelSideBar"
        Me.PanelSideBar.Size = New System.Drawing.Size(219, 706)
        Me.PanelSideBar.TabIndex = 0
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1354, 706)
        Me.Controls.Add(Me.PanelMain)
        Me.Controls.Add(Me.PanelFooter)
        Me.Controls.Add(Me.PanelHeader)
        Me.Controls.Add(Me.PanelSideBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DASHBOARD"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PanelHeader.ResumeLayout(False)
        Me.PanelHeader.PerformLayout()
        Me.PanelTool.ResumeLayout(False)
        CType(Me.btnMenu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelFooter.ResumeLayout(False)
        Me.PanelFooter.PerformLayout()
        Me.PanelMain.ResumeLayout(False)
        Me.PanelStatusBar.ResumeLayout(False)
        Me.PanelStatusBar.PerformLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxSignal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelFormFill.ResumeLayout(False)
        Me.PanelBackGroundImage.ResumeLayout(False)
        Me.PanelBackGroundImage.PerformLayout()
        Me.PanelTitle.ResumeLayout(False)
        Me.PanelTitle.PerformLayout()
        Me.PanelProfile.ResumeLayout(False)
        Me.PanelProfile.PerformLayout()
        CType(Me.PictureBoxProfile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelMenu.ResumeLayout(False)
        Me.PanelMenu.PerformLayout()
        Me.PanelInfo.ResumeLayout(False)
        Me.MenuApp.ResumeLayout(False)
        Me.MenuApp.PerformLayout()
        Me.PanelSideBar.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelHeader As Panel
    Friend WithEvents btnMenu As PictureBox
    Friend WithEvents LabelTitle As Label
    Friend WithEvents PanelFooter As Panel
    Friend WithEvents PanelMain As Panel
    Friend WithEvents Label5 As Label

    Friend WithEvents PanelSparator As Panel
    Friend WithEvents ButtonClose As Button
    Friend WithEvents ButtonMinimize As Button
    Friend WithEvents PanelFormFill As Panel
    Friend WithEvents ProgressBarLoad As ProgressBar
    Friend WithEvents PictureBoxSignal As PictureBox
    Friend WithEvents PanelStatusBar As Panel

    Friend WithEvents LabelSignal As Label
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents LabelDate As Label
    Friend WithEvents TimerMain As Timer
    Friend WithEvents PanelBackGroundImage As Panel

    Friend WithEvents LabelLogo As Label
    Friend WithEvents Button1 As Button

    Friend WithEvents TextBox2 As TextBox

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents btnMessage As Button

    Friend WithEvents PanelTool As Panel

    Friend WithEvents btnLogout As Button
    Friend WithEvents PanelTitle As Panel

    Friend WithEvents LabelRefresh As Label

    Friend WithEvents PanelProfile As Panel

    Friend WithEvents PanelMenu As Panel

    Friend WithEvents PanelInfo As Panel

    Friend WithEvents LabelVersion As Label

    Friend WithEvents MenuApp As MenuStrip

    Friend WithEvents ToolStripMenuItemSitemAdmin As ToolStripMenuItem

    Friend WithEvents UserToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents GroupToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents AuthToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents AuthToolStripMenuItem1 As ToolStripMenuItem

    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents MasterDataToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ProdASPToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ProdPlanToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents SerialNumberToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents SerialNumberFALToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents WorkInstructionToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents BOMToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents MatToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents MaterialPreparelistToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents MSL3ToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents StockCardToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ReturnMaterialNGToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents Prod1ToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents BPRVerificationToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ICMappingToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents RequestICEMMCToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents StencilManagementToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents Prod2ToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents DMSVerificationToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents DMSInlineVerificationToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents LabelBoxCSKDToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents QualityToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ApprovalReqICEMMCToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ScreeningPCBAIDToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents WorksheetPrintToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents WIMasterToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents DelToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents InventoryToolStripMenuItem1 As ToolStripMenuItem

    Friend WithEvents ReportToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents ButtonConnect As Button

    Friend WithEvents ButtonHelp As Button

    Friend WithEvents ButtonSetting As Button

    Friend WithEvents LabelNama As Label

    Friend WithEvents ButtonRefresh As Button

    Friend WithEvents LabelMenu As Label

    Friend WithEvents PictureBoxProfile As PictureBox

    Friend WithEvents LabelUserId As Label

    Friend WithEvents PanelSideBar As Panel
    Friend WithEvents NozzleManagementToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents DinoLiteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimbanganCASToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Button2 As Button
    Friend WithEvents WIMultiMasterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UploadMasterDataToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FGMappingLocatorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LogLabelCSKDToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RepairPCBAToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents RepairCodeMasterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SolderPasteManagementToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FeederManagementToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SolderPasteToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents SolderPasteToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents LabelCompletionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MRPToolStripMenuItem As ToolStripMenuItem

    Friend WithEvents MasterDataToolStripMenuItem1 As ToolStripMenuItem

    Friend WithEvents UploadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RepairInspectionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As ToolStripMenuItem

    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem

    Public Class MyRender
        Inherits ToolStripProfessionalRenderer
        Protected Overloads Overrides Sub OnRenderMenuItemBackground(ByVal e As ToolStripItemRenderEventArgs)
            Try
                e.Item.Padding = New System.Windows.Forms.Padding(0, 7, 0, 7)
                e.Item.Margin = New System.Windows.Forms.Padding(0, 0, 0, 1)
                e.Item.ImageAlign = ContentAlignment.MiddleLeft
                e.Item.TextAlign = ContentAlignment.MiddleLeft
                e.Item.ForeColor = Color.White
                Dim rc As New Rectangle(Point.Empty, e.Item.Size)
                Dim Genap = IIf(e.ToolStrip.Items.IndexOf(e.Item) Mod 2 = 0, True, False)
                Dim c As Color = IIf(e.Item.Selected, System.Drawing.Color.FromArgb(29, 56, 84),
                                     IIf(Genap, Color.FromArgb(51, 133, 166), Color.FromArgb(51, 133, 166)))
                Dim border As Color = IIf(e.Item.Selected, Color.Transparent, Color.Transparent)
                Using brush As New SolidBrush(c)
                    e.Graphics.FillRectangle(brush, rc)
                End Using
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Class
