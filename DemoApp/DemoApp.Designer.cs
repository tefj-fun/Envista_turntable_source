namespace DemoApp
{
    partial class DemoApp
    {
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanelMain = new System.Windows.Forms.TableLayoutPanel();
            this.leftTabs = new System.Windows.Forms.TabControl();
            this.tabWorkflow = new System.Windows.Forms.TabPage();
            this.panelStepsHost = new System.Windows.Forms.Panel();
            this.tableLayoutSteps = new System.Windows.Forms.TableLayoutPanel();
            this.groupStep1 = new System.Windows.Forms.GroupBox();
            this.tableLayoutProject = new System.Windows.Forms.TableLayoutPanel();
            this.TB_ProjectPath = new System.Windows.Forms.TextBox();
            this.BT_LoadProject = new System.Windows.Forms.Button();
            this.groupStep2 = new System.Windows.Forms.GroupBox();
            this.BT_LoadImg = new System.Windows.Forms.Button();
            this.groupStep4 = new System.Windows.Forms.GroupBox();
            this.BT_Detect = new System.Windows.Forms.Button();
            this.groupStep5 = new System.Windows.Forms.GroupBox();
            this.RTB_Info = new System.Windows.Forms.RichTextBox();
            this.groupGallery = new System.Windows.Forms.GroupBox();
            this.flowFrontGallery = new System.Windows.Forms.FlowLayoutPanel();
            this.groupStep7 = new System.Windows.Forms.GroupBox();
            this.DG_Detections = new System.Windows.Forms.DataGridView();
            this.DetectionSequence = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DetectionClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DetectionScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DetectionAngle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DetectionCenter = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DetectionBounds = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutImages = new System.Windows.Forms.TableLayoutPanel();
            this.splitPreview = new System.Windows.Forms.SplitContainer();
            this.groupStep3 = new System.Windows.Forms.GroupBox();
            this.PB_OriginalImage = new System.Windows.Forms.PictureBox();
            this.groupStep6 = new System.Windows.Forms.GroupBox();
            this.tableLayoutFrontPreview = new System.Windows.Forms.TableLayoutPanel();
            this.panelFrontHeader = new System.Windows.Forms.Panel();
            this.LBL_FrontSummary = new System.Windows.Forms.Label();
            this.CHK_ShowOverlay = new System.Windows.Forms.CheckBox();
            this.PB_FrontPreview = new System.Windows.Forms.PictureBox();
            this.panelFrontNav = new System.Windows.Forms.Panel();
            this.BT_FrontPrev = new System.Windows.Forms.Button();
            this.BT_FrontNext = new System.Windows.Forms.Button();
            this.groupDefectLedger = new System.Windows.Forms.GroupBox();
            this.DG_FrontDefects = new System.Windows.Forms.DataGridView();
            this.FrontIndexColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FrontClassColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FrontConfidenceColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FrontAreaColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FrontBoundsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutPanelMain.SuspendLayout();
            this.leftTabs.SuspendLayout();
            this.tabWorkflow.SuspendLayout();
            this.tableLayoutSteps.SuspendLayout();
            this.groupStep1.SuspendLayout();
            this.tableLayoutProject.SuspendLayout();
            this.groupStep2.SuspendLayout();
            this.groupStep4.SuspendLayout();
            this.groupStep5.SuspendLayout();
            this.groupGallery.SuspendLayout();
            this.groupStep7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG_Detections)).BeginInit();
            this.tableLayoutImages.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitPreview)).BeginInit();
            this.splitPreview.Panel1.SuspendLayout();
            this.splitPreview.Panel2.SuspendLayout();
            this.splitPreview.SuspendLayout();
            this.groupStep3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PB_OriginalImage)).BeginInit();
            this.groupStep6.SuspendLayout();
            this.tableLayoutFrontPreview.SuspendLayout();
            this.panelFrontHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PB_FrontPreview)).BeginInit();
            this.panelFrontNav.SuspendLayout();
            this.groupDefectLedger.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG_FrontDefects)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanelMain
            // 
            this.tableLayoutPanelMain.ColumnCount = 2;
            this.tableLayoutPanelMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38F));
            this.tableLayoutPanelMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 62F));
            this.tableLayoutPanelMain.Controls.Add(this.leftTabs, 0, 0);
            this.tableLayoutPanelMain.Controls.Add(this.tableLayoutImages, 1, 0);
            this.tableLayoutPanelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanelMain.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanelMain.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanelMain.Name = "tableLayoutPanelMain";
            this.tableLayoutPanelMain.RowCount = 1;
            this.tableLayoutPanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanelMain.Size = new System.Drawing.Size(1184, 661);
            this.tableLayoutPanelMain.TabIndex = 0;
            // 
            // leftTabs
            // 
            this.leftTabs.Controls.Add(this.tabWorkflow);
            this.leftTabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftTabs.ItemSize = new System.Drawing.Size(120, 28);
            this.leftTabs.Location = new System.Drawing.Point(0, 0);
            this.leftTabs.Margin = new System.Windows.Forms.Padding(0, 0, 6, 0);
            this.leftTabs.Multiline = true;
            this.leftTabs.Name = "leftTabs";
            this.leftTabs.SelectedIndex = 0;
            this.leftTabs.Size = new System.Drawing.Size(449, 661);
            this.leftTabs.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.leftTabs.TabIndex = 0;
            // 
            // tabWorkflow
            // 
            this.tabWorkflow.Controls.Add(this.panelStepsHost);
            this.tabWorkflow.Location = new System.Drawing.Point(4, 32);
            this.tabWorkflow.Margin = new System.Windows.Forms.Padding(0);
            this.tabWorkflow.Name = "tabWorkflow";
            this.tabWorkflow.Padding = new System.Windows.Forms.Padding(0);
            this.tabWorkflow.Size = new System.Drawing.Size(441, 625);
            this.tabWorkflow.TabIndex = 0;
            this.tabWorkflow.Text = "Workflow";
            this.tabWorkflow.UseVisualStyleBackColor = true;
            // 
            // panelStepsHost
            // 
            this.panelStepsHost.AutoScroll = true;
            this.panelStepsHost.AutoScrollMargin = new System.Drawing.Size(0, 16);
            this.panelStepsHost.Controls.Add(this.tableLayoutSteps);
            this.panelStepsHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelStepsHost.Location = new System.Drawing.Point(0, 0);
            this.panelStepsHost.Margin = new System.Windows.Forms.Padding(0);
            this.panelStepsHost.Name = "panelStepsHost";
            this.panelStepsHost.Padding = new System.Windows.Forms.Padding(0);
            this.panelStepsHost.Size = new System.Drawing.Size(441, 625);
            this.panelStepsHost.TabIndex = 0;
            // 
            // tableLayoutSteps
            // 
            this.tableLayoutSteps.AutoSize = true;
            this.tableLayoutSteps.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutSteps.ColumnCount = 1;
            this.tableLayoutSteps.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutSteps.Controls.Add(this.groupStep1, 0, 0);
            this.tableLayoutSteps.Controls.Add(this.groupStep2, 0, 1);
            this.tableLayoutSteps.Controls.Add(this.groupStep4, 0, 2);
            this.tableLayoutSteps.Controls.Add(this.groupGallery, 0, 3);
            this.tableLayoutSteps.Controls.Add(this.groupStep5, 0, 4);
            this.tableLayoutSteps.Controls.Add(this.groupStep7, 0, 5);
            this.tableLayoutSteps.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutSteps.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutSteps.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutSteps.Name = "tableLayoutSteps";
            this.tableLayoutSteps.RowCount = 6;
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutSteps.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutSteps.Size = new System.Drawing.Size(449, 661);
            this.tableLayoutSteps.TabIndex = 0;
            // 
            // groupStep1
            // 
            this.groupStep1.Controls.Add(this.tableLayoutProject);
            this.groupStep1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep1.Location = new System.Drawing.Point(3, 3);
            this.groupStep1.Name = "groupStep1";
            this.groupStep1.Size = new System.Drawing.Size(442, 92);
            this.groupStep1.TabIndex = 0;
            this.groupStep1.TabStop = false;
            this.groupStep1.Text = "1. Select TSP";
            this.groupStep1.Visible = false;
            // 
            // tableLayoutProject
            // 
            this.tableLayoutProject.ColumnCount = 2;
            this.tableLayoutProject.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75F));
            this.tableLayoutProject.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutProject.Controls.Add(this.TB_ProjectPath, 0, 0);
            this.tableLayoutProject.Controls.Add(this.BT_LoadProject, 1, 0);
            this.tableLayoutProject.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutProject.Location = new System.Drawing.Point(3, 18);
            this.tableLayoutProject.Name = "tableLayoutProject";
            this.tableLayoutProject.RowCount = 1;
            this.tableLayoutProject.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutProject.Size = new System.Drawing.Size(436, 71);
            this.tableLayoutProject.TabIndex = 0;
            // 
            // TB_ProjectPath
            // 
            this.TB_ProjectPath.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TB_ProjectPath.Location = new System.Drawing.Point(3, 3);
            this.TB_ProjectPath.Multiline = true;
            this.TB_ProjectPath.Name = "TB_ProjectPath";
            this.TB_ProjectPath.ReadOnly = true;
            this.TB_ProjectPath.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TB_ProjectPath.Size = new System.Drawing.Size(321, 65);
            this.TB_ProjectPath.TabIndex = 0;
            // 
            // BT_LoadProject
            // 
            this.BT_LoadProject.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BT_LoadProject.Location = new System.Drawing.Point(330, 3);
            this.BT_LoadProject.Name = "BT_LoadProject";
            this.BT_LoadProject.Size = new System.Drawing.Size(103, 65);
            this.BT_LoadProject.TabIndex = 1;
            this.BT_LoadProject.Text = "Browse...";
            this.BT_LoadProject.UseVisualStyleBackColor = true;
            this.BT_LoadProject.Click += new System.EventHandler(this.BT_LoadProject_Click);
            // 
            // groupStep2
            // 
            this.groupStep2.Controls.Add(this.BT_LoadImg);
            this.groupStep2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep2.Location = new System.Drawing.Point(3, 101);
            this.groupStep2.Name = "groupStep2";
            this.groupStep2.Size = new System.Drawing.Size(442, 77);
            this.groupStep2.TabIndex = 1;
            this.groupStep2.TabStop = false;
            this.groupStep2.Text = "2. Upload Image";
            this.groupStep2.Visible = false;
            // 
            // BT_LoadImg
            // 
            this.BT_LoadImg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BT_LoadImg.Location = new System.Drawing.Point(3, 18);
            this.BT_LoadImg.Name = "BT_LoadImg";
            this.BT_LoadImg.Size = new System.Drawing.Size(436, 56);
            this.BT_LoadImg.TabIndex = 0;
            this.BT_LoadImg.Text = "Choose Image...";
            this.BT_LoadImg.UseVisualStyleBackColor = true;
            this.BT_LoadImg.Click += new System.EventHandler(this.BT_LoadImg_Click);
            // 
            // groupStep4
            // 
            this.groupStep4.Controls.Add(this.BT_Detect);
            this.groupStep4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep4.Location = new System.Drawing.Point(3, 184);
            this.groupStep4.Name = "groupStep4";
            this.groupStep4.Size = new System.Drawing.Size(442, 77);
            this.groupStep4.TabIndex = 2;
            this.groupStep4.TabStop = false;
            this.groupStep4.Text = "Run Detection";
            // 
            // BT_Detect
            // 
            this.BT_Detect.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BT_Detect.Location = new System.Drawing.Point(3, 18);
            this.BT_Detect.Name = "BT_Detect";
            this.BT_Detect.Size = new System.Drawing.Size(436, 56);
            this.BT_Detect.TabIndex = 0;
            this.BT_Detect.Text = "Run Detection";
            this.BT_Detect.UseVisualStyleBackColor = true;
            this.BT_Detect.Click += new System.EventHandler(this.BT_Detect_Click);
            // 
            // groupGallery
            // 
            this.groupGallery.Controls.Add(this.flowFrontGallery);
            this.groupGallery.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupGallery.Location = new System.Drawing.Point(3, 267);
            this.groupGallery.Name = "groupGallery";
            this.groupGallery.Size = new System.Drawing.Size(442, 296);
            this.groupGallery.TabIndex = 3;
            this.groupGallery.TabStop = false;
            this.groupGallery.Text = "Front Inspections";
            // 
            // flowFrontGallery
            // 
            this.flowFrontGallery.AutoScroll = true;
            this.flowFrontGallery.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowFrontGallery.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowFrontGallery.Location = new System.Drawing.Point(3, 18);
            this.flowFrontGallery.Name = "flowFrontGallery";
            this.flowFrontGallery.Size = new System.Drawing.Size(436, 275);
            this.flowFrontGallery.TabIndex = 0;
            this.flowFrontGallery.WrapContents = false;
            // 
            // groupStep5
            // 
            this.groupStep5.Controls.Add(this.RTB_Info);
            this.groupStep5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep5.Location = new System.Drawing.Point(3, 569);
            this.groupStep5.MinimumSize = new System.Drawing.Size(0, 150);
            this.groupStep5.Name = "groupStep5";
            this.groupStep5.Size = new System.Drawing.Size(442, 150);
            this.groupStep5.TabIndex = 4;
            this.groupStep5.TabStop = false;
            this.groupStep5.Text = "Log";
            // 
            // RTB_Info
            // 
            this.RTB_Info.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RTB_Info.Location = new System.Drawing.Point(3, 18);
            this.RTB_Info.Name = "RTB_Info";
            this.RTB_Info.Size = new System.Drawing.Size(436, 129);
            this.RTB_Info.TabIndex = 0;
            this.RTB_Info.Text = "";
            // 
            // groupStep7
            // 
            this.groupStep7.Controls.Add(this.DG_Detections);
            this.groupStep7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep7.Location = new System.Drawing.Point(3, 428);
            this.groupStep7.Name = "groupStep7";
            this.groupStep7.Size = new System.Drawing.Size(442, 230);
            this.groupStep7.TabIndex = 4;
            this.groupStep7.TabStop = false;
            this.groupStep7.Text = "7. Detection Results";
            // 
            // DG_Detections
            // 
            this.DG_Detections.AllowUserToAddRows = false;
            this.DG_Detections.AllowUserToDeleteRows = false;
            this.DG_Detections.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DG_Detections.BackgroundColor = System.Drawing.SystemColors.Window;
            this.DG_Detections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG_Detections.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DetectionSequence,
            this.DetectionClass,
            this.DetectionScore,
            this.DetectionAngle,
            this.DetectionCenter,
            this.DetectionBounds});
            this.DG_Detections.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DG_Detections.Location = new System.Drawing.Point(3, 18);
            this.DG_Detections.MultiSelect = false;
            this.DG_Detections.Name = "DG_Detections";
            this.DG_Detections.ReadOnly = true;
            this.DG_Detections.RowHeadersVisible = false;
            this.DG_Detections.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DG_Detections.Size = new System.Drawing.Size(436, 209);
            this.DG_Detections.TabIndex = 0;
            this.DG_Detections.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DG_Detections_CellMouseClick);
            // 
            // DetectionSequence
            // 
            this.DetectionSequence.HeaderText = "#";
            this.DetectionSequence.Name = "DetectionSequence";
            this.DetectionSequence.ReadOnly = true;
            // 
            // DetectionClass
            // 
            this.DetectionClass.HeaderText = "Class";
            this.DetectionClass.Name = "DetectionClass";
            this.DetectionClass.ReadOnly = true;
            // 
            // DetectionScore
            // 
            this.DetectionScore.HeaderText = "Score";
            this.DetectionScore.Name = "DetectionScore";
            this.DetectionScore.ReadOnly = true;
            // 
            // DetectionAngle
            // 
            this.DetectionAngle.HeaderText = "Angle (deg)";
            this.DetectionAngle.Name = "DetectionAngle";
            this.DetectionAngle.ReadOnly = true;
            // 
            // DetectionCenter
            // 
            this.DetectionCenter.HeaderText = "Center (x,y)";
            this.DetectionCenter.Name = "DetectionCenter";
            this.DetectionCenter.ReadOnly = true;
            // 
            // DetectionBounds
            // 
            this.DetectionBounds.HeaderText = "Bounds (x,y,w,h)";
            this.DetectionBounds.Name = "DetectionBounds";
            this.DetectionBounds.ReadOnly = true;
            // 
            // tableLayoutImages
            // 
            this.tableLayoutImages.ColumnCount = 1;
            this.tableLayoutImages.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutImages.Controls.Add(this.splitPreview, 0, 0);
            this.tableLayoutImages.Controls.Add(this.groupDefectLedger, 0, 1);
            this.tableLayoutImages.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutImages.Location = new System.Drawing.Point(454, 0);
            this.tableLayoutImages.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutImages.Name = "tableLayoutImages";
            this.tableLayoutImages.RowCount = 2;
            this.tableLayoutImages.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tableLayoutImages.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tableLayoutImages.Size = new System.Drawing.Size(730, 661);
            this.tableLayoutImages.TabIndex = 1;
            // 
            // splitPreview
            // 
            this.splitPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitPreview.Location = new System.Drawing.Point(3, 3);
            this.splitPreview.Name = "splitPreview";
            this.splitPreview.Panel1.Controls.Add(this.groupStep3);
            this.splitPreview.Panel2.Controls.Add(this.groupStep6);
            this.splitPreview.Size = new System.Drawing.Size(724, 424);
            this.splitPreview.SplitterDistance = 360;
            this.splitPreview.TabIndex = 2;
            // 
            // groupStep3
            // 
            this.groupStep3.Controls.Add(this.PB_OriginalImage);
            this.groupStep3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep3.Location = new System.Drawing.Point(0, 0);
            this.groupStep3.Name = "groupStep3";
            this.groupStep3.Size = new System.Drawing.Size(360, 424);
            this.groupStep3.TabIndex = 0;
            this.groupStep3.TabStop = false;
            this.groupStep3.Text = "3. Attachment Overview";
            // 
            // PB_OriginalImage
            // 
            this.PB_OriginalImage.BackColor = System.Drawing.Color.Black;
            this.PB_OriginalImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PB_OriginalImage.Location = new System.Drawing.Point(3, 18);
            this.PB_OriginalImage.Name = "PB_OriginalImage";
            this.PB_OriginalImage.Size = new System.Drawing.Size(354, 403);
            this.PB_OriginalImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PB_OriginalImage.TabIndex = 0;
            this.PB_OriginalImage.TabStop = false;
            // 
            // groupStep6
            // 
            this.groupStep6.Controls.Add(this.tableLayoutFrontPreview);
            this.groupStep6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupStep6.Location = new System.Drawing.Point(0, 0);
            this.groupStep6.Name = "groupStep6";
            this.groupStep6.Size = new System.Drawing.Size(360, 424);
            this.groupStep6.TabIndex = 1;
            this.groupStep6.TabStop = false;
            this.groupStep6.Text = "4. Front Inspection";
            // 
            // tableLayoutFrontPreview
            // 
            this.tableLayoutFrontPreview.ColumnCount = 1;
            this.tableLayoutFrontPreview.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutFrontPreview.Controls.Add(this.panelFrontHeader, 0, 0);
            this.tableLayoutFrontPreview.Controls.Add(this.PB_FrontPreview, 0, 1);
            this.tableLayoutFrontPreview.Controls.Add(this.panelFrontNav, 0, 2);
            this.tableLayoutFrontPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutFrontPreview.Location = new System.Drawing.Point(3, 18);
            this.tableLayoutFrontPreview.Name = "tableLayoutFrontPreview";
            this.tableLayoutFrontPreview.RowCount = 3;
            this.tableLayoutFrontPreview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutFrontPreview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutFrontPreview.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutFrontPreview.Size = new System.Drawing.Size(354, 403);
            this.tableLayoutFrontPreview.TabIndex = 0;
            // 
            // panelFrontHeader
            // 
            this.panelFrontHeader.Controls.Add(this.LBL_FrontSummary);
            this.panelFrontHeader.Controls.Add(this.CHK_ShowOverlay);
            this.panelFrontHeader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFrontHeader.Location = new System.Drawing.Point(3, 3);
            this.panelFrontHeader.Name = "panelFrontHeader";
            this.panelFrontHeader.Size = new System.Drawing.Size(348, 32);
            this.panelFrontHeader.TabIndex = 0;
            // 
            // LBL_FrontSummary
            // 
            this.LBL_FrontSummary.AutoSize = true;
            this.LBL_FrontSummary.Location = new System.Drawing.Point(3, 8);
            this.LBL_FrontSummary.Name = "LBL_FrontSummary";
            this.LBL_FrontSummary.Size = new System.Drawing.Size(140, 12);
            this.LBL_FrontSummary.TabIndex = 1;
            this.LBL_FrontSummary.Text = "No inspection selected.";
            // 
            // CHK_ShowOverlay
            // 
            this.CHK_ShowOverlay.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CHK_ShowOverlay.AutoSize = true;
            this.CHK_ShowOverlay.Checked = true;
            this.CHK_ShowOverlay.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CHK_ShowOverlay.Location = new System.Drawing.Point(240, 6);
            this.CHK_ShowOverlay.Name = "CHK_ShowOverlay";
            this.CHK_ShowOverlay.Size = new System.Drawing.Size(99, 16);
            this.CHK_ShowOverlay.TabIndex = 0;
            this.CHK_ShowOverlay.Text = "Show overlay";
            this.CHK_ShowOverlay.UseVisualStyleBackColor = true;
            this.CHK_ShowOverlay.CheckedChanged += new System.EventHandler(this.CHK_ShowOverlay_CheckedChanged);
            // 
            // PB_FrontPreview
            // 
            this.PB_FrontPreview.BackColor = System.Drawing.Color.Black;
            this.PB_FrontPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PB_FrontPreview.Location = new System.Drawing.Point(3, 41);
            this.PB_FrontPreview.Name = "PB_FrontPreview";
            this.PB_FrontPreview.Size = new System.Drawing.Size(348, 329);
            this.PB_FrontPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PB_FrontPreview.TabIndex = 1;
            this.PB_FrontPreview.TabStop = false;
            // 
            // panelFrontNav
            // 
            this.panelFrontNav.Controls.Add(this.BT_FrontPrev);
            this.panelFrontNav.Controls.Add(this.BT_FrontNext);
            this.panelFrontNav.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFrontNav.Location = new System.Drawing.Point(3, 376);
            this.panelFrontNav.Name = "panelFrontNav";
            this.panelFrontNav.Size = new System.Drawing.Size(348, 24);
            this.panelFrontNav.TabIndex = 2;
            // 
            // BT_FrontPrev
            // 
            this.BT_FrontPrev.Dock = System.Windows.Forms.DockStyle.Left;
            this.BT_FrontPrev.Location = new System.Drawing.Point(0, 0);
            this.BT_FrontPrev.Name = "BT_FrontPrev";
            this.BT_FrontPrev.Size = new System.Drawing.Size(75, 24);
            this.BT_FrontPrev.TabIndex = 0;
            this.BT_FrontPrev.Text = "Previous";
            this.BT_FrontPrev.UseVisualStyleBackColor = true;
            this.BT_FrontPrev.Click += new System.EventHandler(this.BT_FrontPrev_Click);
            // 
            // BT_FrontNext
            // 
            this.BT_FrontNext.Dock = System.Windows.Forms.DockStyle.Right;
            this.BT_FrontNext.Location = new System.Drawing.Point(273, 0);
            this.BT_FrontNext.Name = "BT_FrontNext";
            this.BT_FrontNext.Size = new System.Drawing.Size(75, 24);
            this.BT_FrontNext.TabIndex = 1;
            this.BT_FrontNext.Text = "Next";
            this.BT_FrontNext.UseVisualStyleBackColor = true;
            this.BT_FrontNext.Click += new System.EventHandler(this.BT_FrontNext_Click);
            // 
            // groupDefectLedger
            // 
            this.groupDefectLedger.Controls.Add(this.DG_FrontDefects);
            this.groupDefectLedger.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupDefectLedger.Location = new System.Drawing.Point(3, 433);
            this.groupDefectLedger.Name = "groupDefectLedger";
            this.groupDefectLedger.Size = new System.Drawing.Size(724, 225);
            this.groupDefectLedger.TabIndex = 3;
            this.groupDefectLedger.TabStop = false;
            this.groupDefectLedger.Text = "Defect Ledger";
            // 
            // DG_FrontDefects
            // 
            this.DG_FrontDefects.AllowUserToAddRows = false;
            this.DG_FrontDefects.AllowUserToDeleteRows = false;
            this.DG_FrontDefects.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DG_FrontDefects.BackgroundColor = System.Drawing.SystemColors.Window;
            this.DG_FrontDefects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG_FrontDefects.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FrontIndexColumn,
            this.FrontClassColumn,
            this.FrontConfidenceColumn,
            this.FrontAreaColumn,
            this.FrontBoundsColumn});
            this.DG_FrontDefects.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DG_FrontDefects.Location = new System.Drawing.Point(3, 18);
            this.DG_FrontDefects.MultiSelect = false;
            this.DG_FrontDefects.Name = "DG_FrontDefects";
            this.DG_FrontDefects.ReadOnly = true;
            this.DG_FrontDefects.RowHeadersVisible = false;
            this.DG_FrontDefects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DG_FrontDefects.Size = new System.Drawing.Size(718, 204);
            this.DG_FrontDefects.TabIndex = 0;
            // 
            // FrontIndexColumn
            // 
            this.FrontIndexColumn.HeaderText = "Index";
            this.FrontIndexColumn.Name = "FrontIndexColumn";
            this.FrontIndexColumn.ReadOnly = true;
            // 
            // FrontClassColumn
            // 
            this.FrontClassColumn.HeaderText = "Class";
            this.FrontClassColumn.Name = "FrontClassColumn";
            this.FrontClassColumn.ReadOnly = true;
            // 
            // FrontConfidenceColumn
            // 
            this.FrontConfidenceColumn.HeaderText = "Confidence";
            this.FrontConfidenceColumn.Name = "FrontConfidenceColumn";
            this.FrontConfidenceColumn.ReadOnly = true;
            // 
            // FrontAreaColumn
            // 
            this.FrontAreaColumn.HeaderText = "Area";
            this.FrontAreaColumn.Name = "FrontAreaColumn";
            this.FrontAreaColumn.ReadOnly = true;
            // 
            // FrontBoundsColumn
            // 
            this.FrontBoundsColumn.HeaderText = "Bounds (x,y,w,h)";
            this.FrontBoundsColumn.Name = "FrontBoundsColumn";
            this.FrontBoundsColumn.ReadOnly = true;
            // 
            // DemoApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1184, 661);
            this.Controls.Add(this.tableLayoutPanelMain);
            this.Name = "DemoApp";
            this.Text = "SolVision Demo";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DemoApp_FormClosing);
            this.tableLayoutPanelMain.ResumeLayout(false);
            this.leftTabs.ResumeLayout(false);
            this.tabWorkflow.ResumeLayout(false);
            this.panelStepsHost.ResumeLayout(false);
            this.panelStepsHost.PerformLayout();
            this.tableLayoutSteps.ResumeLayout(false);
            this.groupStep1.ResumeLayout(false);
            this.tableLayoutProject.ResumeLayout(false);
            this.tableLayoutProject.PerformLayout();
            this.groupStep2.ResumeLayout(false);
            this.groupStep4.ResumeLayout(false);
            this.groupGallery.ResumeLayout(false);
            this.flowFrontGallery.ResumeLayout(false);
            this.flowFrontGallery.PerformLayout();
            this.groupStep5.ResumeLayout(false);
            this.groupStep7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DG_Detections)).EndInit();
            this.tableLayoutImages.ResumeLayout(false);
            this.splitPreview.Panel1.ResumeLayout(false);
            this.splitPreview.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitPreview)).EndInit();
            this.splitPreview.ResumeLayout(false);
            this.groupStep3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PB_OriginalImage)).EndInit();
            this.groupStep6.ResumeLayout(false);
            this.tableLayoutFrontPreview.ResumeLayout(false);
            this.tableLayoutFrontPreview.PerformLayout();
            this.panelFrontHeader.ResumeLayout(false);
            this.panelFrontHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PB_FrontPreview)).EndInit();
            this.panelFrontNav.ResumeLayout(false);
            this.groupDefectLedger.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DG_FrontDefects)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelMain;
        private System.Windows.Forms.Panel panelStepsHost;
        private System.Windows.Forms.TableLayoutPanel tableLayoutSteps;
        private System.Windows.Forms.GroupBox groupStep1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutProject;
        private System.Windows.Forms.TextBox TB_ProjectPath;
        private System.Windows.Forms.Button BT_LoadProject;
        private System.Windows.Forms.GroupBox groupStep2;
        private System.Windows.Forms.Button BT_LoadImg;
        private System.Windows.Forms.GroupBox groupStep4;
        private System.Windows.Forms.Button BT_Detect;
        private System.Windows.Forms.GroupBox groupStep5;
        private System.Windows.Forms.RichTextBox RTB_Info;
        private System.Windows.Forms.GroupBox groupGallery;
        private System.Windows.Forms.FlowLayoutPanel flowFrontGallery;
        private System.Windows.Forms.GroupBox groupStep7;
        private System.Windows.Forms.DataGridView DG_Detections;
        private System.Windows.Forms.TableLayoutPanel tableLayoutImages;
        private System.Windows.Forms.GroupBox groupStep3;
        private System.Windows.Forms.PictureBox PB_OriginalImage;
        private System.Windows.Forms.GroupBox groupStep6;
        private System.Windows.Forms.TableLayoutPanel tableLayoutFrontPreview;
        private System.Windows.Forms.Panel panelFrontHeader;
        private System.Windows.Forms.Label LBL_FrontSummary;
        private System.Windows.Forms.CheckBox CHK_ShowOverlay;
        private System.Windows.Forms.PictureBox PB_FrontPreview;
        private System.Windows.Forms.Panel panelFrontNav;
        private System.Windows.Forms.Button BT_FrontPrev;
        private System.Windows.Forms.Button BT_FrontNext;
        private System.Windows.Forms.SplitContainer splitPreview;
        private System.Windows.Forms.GroupBox groupDefectLedger;
        private System.Windows.Forms.DataGridView DG_FrontDefects;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionSequence;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionAngle;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionCenter;
        private System.Windows.Forms.DataGridViewTextBoxColumn DetectionBounds;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrontIndexColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrontClassColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrontConfidenceColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrontAreaColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrontBoundsColumn;
    }
}
