using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using SolVision;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using MVSDK_Net;

namespace DemoApp
{
    public partial class DemoApp : Form
    {
        private enum LogStatus
        {
            Info,
            Success,
            Warning,
            Error,
            Progress
        }

        private SolVision.TaskProcess SolDL;
        private DetectImg currentImage;
        private List<string> classNameList;
        private string loadedProjectPath = string.Empty;
        private RuleGroupNode logicRoot;
        private bool suppressLogicEvents;
        private readonly Color statusPassBackground = Color.FromArgb(232, 245, 233);
        private readonly Color statusPassForeground = Color.FromArgb(27, 94, 32);
        private readonly Color statusFailBackground = Color.FromArgb(255, 235, 238);
        private readonly Color statusFailForeground = Color.FromArgb(178, 34, 34);
        private readonly Color statusNeutralBackground = Color.FromArgb(245, 247, 250);
        private readonly Color statusNeutralForeground = Color.FromArgb(94, 102, 112);
        private TabControl leftTabs;
        private TabPage tabWorkflow;
        private TabPage tabLogicBuilder;
        private TableLayoutPanel workflowLayout;
        private Label LBL_WorkflowStatus;
        private TableLayoutPanel logicRootLayout;
        private Panel logicHeaderPanel;
        private Label labelLogicHeader;
        private Button BT_LogicHelp;
        private ToolTip toolTip;
        private Panel loadingIndicatorPanel;
        private ProgressBar loadingSpinner;
        private int loadingRequestCount;
        private GroupBox groupLogic;
        private TableLayoutPanel tableLayoutLogic;
        private TreeView TV_Logic;
        private Panel panelLogicEditor;
        private TableLayoutPanel tableLayoutLogicEditor;
        private Label labelLogicGroup;
        private ComboBox CB_GroupOperator;
        private Label labelLogicField;
        private ComboBox CB_Field;
        private Label labelLogicOperator;
        private ComboBox CB_Operator;
        private Label labelLogicValue;
        private ComboBox CB_Value;
        private TableLayoutPanel tableLayoutLogicButtons;
        private Button BT_AddRule;
        private Button BT_AddGroup;
        private Button BT_RemoveNode;
        private Label LBL_ResultStatus;

        private ModelContext attachmentContext;
        private ModelContext defectContext;

        private TabPage tabInitialize;
        private TableLayoutPanel initLayout;
        private GroupBox groupInitAttachment;
        private GroupBox groupInitDefect;
        private TextBox TB_InitAttachmentPath;
        private Button BT_InitAttachmentBrowse;
        private Button BT_InitAttachmentLoad;
        private Label LBL_InitAttachmentStatus;
        private TextBox TB_InitDefectPath;
        private Button BT_InitDefectBrowse;
        private Button BT_InitDefectLoad;
        private Label LBL_InitDefectStatus;
        private Label LBL_InitSummary;

        private GroupBox groupInitCameras;
        private Button BT_CameraRefresh;
        private ComboBox CB_TopCameraSelect;
        private ComboBox CB_FrontCameraSelect;
        private Button BT_TopCameraConnect;
        private Button BT_FrontCameraConnect;
        private Button BT_TopCameraCapture;
        private Button BT_FrontCameraCapture;
        private Label LBL_TopCameraStatus;
        private Label LBL_FrontCameraStatus;

        private TurntableController turntableController;
        private GroupBox groupInitTurntable;
        private ComboBox CB_TurntablePort;
        private Button BT_TurntableRefresh;
        private Button BT_TurntableConnect;
        private Button BT_TurntableHome;
        private Label LBL_TurntableStatus;

        private CameraContext topCameraContext;
        private CameraContext frontCameraContext;
        private readonly List<CameraDeviceInfo> cameraDeviceCache = new List<CameraDeviceInfo>();
        private CancellationTokenSource captureSequenceCts;
        private bool showFrontOverlay = true;
        private int selectedFrontSequence = -1;

        public DemoApp()
        {
            InitializeComponent();
            classNameList = new List<string>();
            topCameraContext = new CameraContext(CameraRole.Top);
            frontCameraContext = new CameraContext(CameraRole.Front);
            turntableController = new TurntableController();
            turntableController.MessageReceived += TurntableController_MessageReceived;
            BuildLogicBuilderUI();
            groupStep3.Text = "Attachment Overview";
            groupStep6.Text = "Front Inspection";
            if (groupStep7 != null)
            {
                groupStep7.Text = "Top Detections";
            }
            if (groupGallery != null)
            {
                groupGallery.Text = "Front Inspections";
            }
            if (groupDefectLedger != null)
            {
                groupDefectLedger.Text = "Defect Ledger";
            }
            InitializeLoadingIndicator();
            InitialSolVision();
            ApplyTheme();
            InitializeLogicBuilder();
        }



private void DemoApp_FormClosing(object sender, FormClosingEventArgs e)

{

    CancelAttachmentSequence();

    if (defectContext?.Process != null)
    {
        defectContext.Process.UpdateClassNamesEvent -= UpdateClassNames;
    }
    if (attachmentContext?.Process != null)
    {
        attachmentContext.Process.UpdateClassNamesEvent -= UpdateClassNames;
    }

    attachmentContext?.Dispose();

    defectContext?.Dispose();

    topCameraContext?.Dispose();

    frontCameraContext?.Dispose();



    if (turntableController != null)

    {

        turntableController.MessageReceived -= TurntableController_MessageReceived;

        turntableController.Dispose();

    }



    currentImage?.Dispose();

}



        public void InitialSolVision()
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application();
            }

            attachmentContext = CreateModelContext("Attachment");
            defectContext = CreateModelContext("Defects");

            SolDL = attachmentContext.Process;
            if (defectContext?.Process != null)
            {
                defectContext.Process.UpdateClassNamesEvent += UpdateClassNames;
            }

            BindModelContextUI();
        }

        private ModelContext CreateModelContext(string name)
        {
            var process = new SolVision.TaskProcess(ExecuteType.Dll);
            ElementHost.EnableModelessKeyboardInterop(process);
            process.Visibility = System.Windows.Visibility.Hidden;
            return new ModelContext(name, process);
        }

        private void BindModelContextUI()
        {
            if (attachmentContext != null)
            {
                attachmentContext.PathDisplay = TB_InitAttachmentPath;
                attachmentContext.StatusDisplay = LBL_InitAttachmentStatus;
                attachmentContext.BrowseButton = BT_InitAttachmentBrowse;
                attachmentContext.LoadButton = BT_InitAttachmentLoad;
                UpdateModelStatus(attachmentContext, "Not loaded.", statusNeutralBackground, statusNeutralForeground);
                ToggleInitButtons(attachmentContext, true);
            }

            if (defectContext != null)
            {
                defectContext.PathDisplay = TB_InitDefectPath;
                defectContext.StatusDisplay = LBL_InitDefectStatus;
                defectContext.BrowseButton = BT_InitDefectBrowse;
                defectContext.LoadButton = BT_InitDefectLoad;
                UpdateModelStatus(defectContext, "Not loaded.", statusNeutralBackground, statusNeutralForeground);
                ToggleInitButtons(defectContext, true);
            }

            PopulateTurntablePorts();
        }

        private async void BT_LoadProject_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "SolVision Project (*.tsp)|*.tsp";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    await LoadProjectAsync(attachmentContext, dlg.FileName, updateWorkflowPath: true);
                }
            }
        }

        private void BT_LoadImg_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Image (*.png;*.jpg;*.bmp;*.tif)|*.png;*.jpg;*.bmp;*.tif";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (!File.Exists(dlg.FileName))
                    {
                        outToLog("Selected image file does not exist.", LogStatus.Error);
                        return;
                    }

                    outToLog("Loading image...", LogStatus.Progress);

                    currentImage?.Dispose();
                    currentImage = new DetectImg
                    {
                        OriImgPath = dlg.FileName,
                        OriImg = CvInvoke.Imread(dlg.FileName, ImreadModes.ColorBgr),
                        ObjList = new List<ObjectInfo>(),
                        AttachmentPoints = new List<AttachmentPointInfo>(),
                        FrontInspections = new List<FrontInspectionResult>(),
                        FrontInspectionComplete = false
                    };
                    currentImage.HasResults = false;

                    DisplayOriginalImage();
                    ClearDetectionVisuals();

                    outToLog($"Image loaded: {Path.GetFileName(dlg.FileName)}", LogStatus.Success);
                }
            }
        }

        private async Task LoadProjectAsync(ModelContext context, string filePath, bool updateWorkflowPath)
        {
            if (context == null)
            {
                outToLog("Model context is not ready.", LogStatus.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                UpdateModelStatus(context, "Invalid project path.", statusFailBackground, statusFailForeground);
                outToLog($"[{context.Name}] Project path is empty.", LogStatus.Error);
                return;
            }

            if (!File.Exists(filePath))
            {
                UpdateModelStatus(context, "Project file not found.", statusFailBackground, statusFailForeground);
                outToLog($"[{context.Name}] Project file not found: {filePath}", LogStatus.Error);
                return;
            }

            string fileName = Path.GetFileName(filePath);
            outToLog($"[{context.Name}] Loading project: {fileName}...", LogStatus.Progress);
            UpdateModelStatus(context, $"Loading {fileName}...", statusNeutralBackground, statusNeutralForeground);

            SetWorkflowControlsEnabled(false);
            ToggleInitButtons(context, false);
            ShowLoadingIndicator($"Loading project: {fileName}...");
            UseWaitCursor = true;

            string errorMessage = null;
            bool success = false;

            try
            {
                await Task.Run(() =>
                {
                    SolVision.TaskProcess.SetLoadingInThread(true);
                    try
                    {
                        var dispatcher = System.Windows.Application.Current?.Dispatcher;
                        if (dispatcher != null)
                        {
                            dispatcher.Invoke(() => System.Windows.Application.Current.MainWindow = context.Process);
                        }

                        context.Process.LoadProject(filePath);
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                    }
                    finally
                    {
                        SolVision.TaskProcess.SetLoadingInThread(false);
                    }
                });
            }
            finally
            {
                HideLoadingIndicator();
                SetWorkflowControlsEnabled(true);
                ToggleInitButtons(context, true);
                UseWaitCursor = false;
            }

            if (!success)
            {
                string failureMsg = $"[{context.Name}] Project load failed: {errorMessage ?? "Unknown error"}";
                UpdateModelStatus(context, failureMsg, statusFailBackground, statusFailForeground);
                outToLog(failureMsg, LogStatus.Error);
                context.IsLoaded = false;
                context.LastError = errorMessage;
                UpdateInitSummary();
                return;
            }

            context.LoadedPath = filePath;
            context.IsLoaded = true;
            context.LastError = null;

            if (context.PathDisplay != null)
            {
                context.PathDisplay.Text = filePath;
            }

            UpdateModelStatus(context, $"Loaded {fileName}", statusPassBackground, statusPassForeground);
            outToLog($"[{context.Name}] Project loaded: {fileName}", LogStatus.Success);

            if (context == attachmentContext && updateWorkflowPath)
            {
                loadedProjectPath = filePath;
                TB_ProjectPath.Text = filePath;
                SolDL = context.Process;
            }

            if (context == defectContext)
            {
                RefreshDefectClassNamesFromModel();
            }

            UpdateInitSummary();
            SetLogicStatusNeutral("Awaiting defect inspection.");
            UpdateLogicEvaluation();
        }

        private void BT_InitAttachmentBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "SolVision Project (*.tsp)|*.tsp";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    TB_InitAttachmentPath.Text = dlg.FileName;
                    ToggleInitButtons(attachmentContext, true);
                    UpdateModelStatus(attachmentContext, "Ready to load.", statusNeutralBackground, statusNeutralForeground);
                }
            }
        }

        private void BT_InitDefectBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "SolVision Project (*.tsp)|*.tsp";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    TB_InitDefectPath.Text = dlg.FileName;
                    ToggleInitButtons(defectContext, true);
                    UpdateModelStatus(defectContext, "Ready to load.", statusNeutralBackground, statusNeutralForeground);
                }
            }
        }

        private async void BT_InitAttachmentLoad_Click(object sender, EventArgs e)
        {
            if (attachmentContext == null)
            {
                outToLog("Attachment context is not ready.", LogStatus.Error);
                return;
            }

            string path = TB_InitAttachmentPath.Text;
            if (!File.Exists(path))
            {
                UpdateModelStatus(attachmentContext, "Project file not found.", statusFailBackground, statusFailForeground);
                outToLog($"[Attachment] Project file not found: {path}", LogStatus.Error);
                return;
            }

            await LoadProjectAsync(attachmentContext, path, updateWorkflowPath: true);
        }

        private async void BT_InitDefectLoad_Click(object sender, EventArgs e)
        {
            if (defectContext == null)
            {
                outToLog("Defect context is not ready.", LogStatus.Error);
                return;
            }

            string path = TB_InitDefectPath.Text;
            if (!File.Exists(path))
            {
                UpdateModelStatus(defectContext, "Project file not found.", statusFailBackground, statusFailForeground);
                outToLog($"[Defect] Project file not found: {path}", LogStatus.Error);
                return;
            }

            await LoadProjectAsync(defectContext, path, updateWorkflowPath: false);
        }


        private void SetWorkflowControlsEnabled(bool enabled)
        {
            BT_LoadProject.Enabled = enabled;
            BT_LoadImg.Enabled = enabled;
            BT_Detect.Enabled = enabled;
        }

        private void ToggleInitButtons(ModelContext context, bool enabled)
        {
            if (context?.BrowseButton != null)
            {
                context.BrowseButton.Enabled = enabled;
            }

            if (context?.LoadButton != null)
            {
                bool hasPath = context.PathDisplay != null && !string.IsNullOrWhiteSpace(context.PathDisplay.Text);
                context.LoadButton.Enabled = enabled && hasPath;
            }
        }

        private TableLayoutPanel CreateCameraRoleLayout(
            CameraContext context,
            string headerText,
            out ComboBox selector,
            out Button connectButton,
            out Button captureButton,
            out Label detailLabel,
            out Label statusLabel)
        {
            TableLayoutPanel layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(6, 0, 6, 0),
                RowCount = 5
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));

            Label header = new Label
            {
                Text = headerText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Padding = new Padding(0, 4, 0, 4)
            };
            layout.Controls.Add(header, 0, 0);

            detailLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5F),
                Padding = new Padding(0, 0, 0, 4)
            };
            layout.Controls.Add(detailLabel, 0, 1);

            selector = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FormattingEnabled = true,
                DisplayMember = nameof(CameraDeviceInfo.DisplayName),
                ValueMember = nameof(CameraDeviceInfo.SerialNumber)
            };
            layout.Controls.Add(selector, 0, 2);

            TableLayoutPanel buttonRow = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            buttonRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            buttonRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            connectButton = new Button
            {
                Text = "Connect",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            buttonRow.Controls.Add(connectButton, 0, 0);

            captureButton = new Button
            {
                Text = "Capture Preview",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            buttonRow.Controls.Add(captureButton, 1, 0);

            layout.Controls.Add(buttonRow, 0, 3);

            statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 6, 8, 6)
            };
            layout.Controls.Add(statusLabel, 0, 4);

            return layout;
        }

        private void BT_CameraRefresh_Click(object sender, EventArgs e)
        {
            RefreshCameraList();
        }

        private void RefreshCameraList()
        {
            try
            {
                cameraDeviceCache.Clear();
                topCameraContext?.ClearPreview();
                frontCameraContext?.ClearPreview();
                ClearWorkflowPreview(CameraRole.Top);
                ClearWorkflowPreview(CameraRole.Front);

                IMVDefine.IMV_DeviceList deviceList = new IMVDefine.IMV_DeviceList();
                int result = MyCamera.IMV_EnumDevices(ref deviceList, (uint)IMVDefine.IMV_EInterfaceType.interfaceTypeAll);

                if (result == IMVDefine.IMV_OK && deviceList.nDevNum > 0)
                {
                    int structSize = Marshal.SizeOf(typeof(IMVDefine.IMV_DeviceInfo));
                    for (int i = 0; i < deviceList.nDevNum; i++)
                    {
                        IntPtr infoPtr = deviceList.pDevInfo + structSize * i;
                        IMVDefine.IMV_DeviceInfo deviceInfo = (IMVDefine.IMV_DeviceInfo)Marshal.PtrToStructure(
                            infoPtr,
                            typeof(IMVDefine.IMV_DeviceInfo));

                        CameraDeviceInfo device = new CameraDeviceInfo(
                            i,
                            deviceInfo.cameraName,
                            deviceInfo.serialNumber,
                            deviceInfo.vendorName,
                            deviceInfo.nCameraType,
                            deviceInfo.nInterfaceType);

                        cameraDeviceCache.Add(device);
                        outToLog($"[Camera] Detected {device.DisplayName}", LogStatus.Info);
                    }

                    outToLog($"[Camera] Found {cameraDeviceCache.Count} device(s).", LogStatus.Info);
                }
                else
                {
                    outToLog("[Camera] No devices detected.", LogStatus.Warning);
                }
            }
            catch (Exception ex)
            {
                outToLog($"[Camera] Enumeration failed: {ex.Message}", LogStatus.Error);
            }
            finally
            {
                ApplyCameraListToSelectors();
            }
        }

        private void ApplyCameraListToSelectors()
        {
            ApplyCameraListToSelector(topCameraContext);
            ApplyCameraListToSelector(frontCameraContext);
        }

        private void ApplyCameraListToSelector(CameraContext context)
        {
            if (context?.Selector == null)
            {
                return;
            }

            object previousSelection = context.IsConnected
                ? context.ConnectedDevice
                : context.Selector.SelectedItem;

            context.Selector.BeginUpdate();
            context.Selector.DisplayMember = nameof(CameraDeviceInfo.DisplayName);
            context.Selector.Items.Clear();

            foreach (CameraDeviceInfo device in cameraDeviceCache)
            {
                context.Selector.Items.Add(device);
            }

            if (context.IsConnected)
            {
                CameraDeviceInfo match = cameraDeviceCache.FirstOrDefault(d =>
                    string.Equals(d.SerialNumber, context.ConnectedDevice.SerialNumber, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    context.Selector.SelectedItem = match;
                }
                else
                {
                    outToLog($"[Camera] {context.RoleName} device disconnected. Closing handle.", LogStatus.Warning);
                    context.Disconnect();
                    context.Selector.SelectedIndex = -1;
                    context.ClearPreview();
                    ClearWorkflowPreview(context.Role);
                }
            }
            else if (previousSelection is CameraDeviceInfo previous)
            {
                CameraDeviceInfo match = cameraDeviceCache.FirstOrDefault(d =>
                    string.Equals(d.SerialNumber, previous.SerialNumber, StringComparison.OrdinalIgnoreCase));
                context.Selector.SelectedItem = match;
            }
            else
            {
                context.Selector.SelectedIndex = cameraDeviceCache.Count > 0 ? 0 : -1;
            }

            context.Selector.EndUpdate();
            RefreshCameraUiState(context);
        }

        private void UpdateCameraSelectionChanged(CameraContext context)
        {
            RefreshCameraUiState(context);
        }

        private void RefreshCameraUiState(CameraContext context)
        {
            if (context == null)
            {
                return;
            }

            bool hasSelection = context.Selector?.SelectedItem is CameraDeviceInfo;

            if (context.ConnectButton != null)
            {
                context.ConnectButton.Text = context.IsConnected ? "Disconnect" : "Connect Camera";
                context.ConnectButton.Enabled = context.IsConnected || hasSelection;
            }

            if (context.CaptureButton != null)
            {
                context.CaptureButton.Enabled = context.IsConnected;
            }

            if (context.DetailLabel != null)
            {
                if (context.IsConnected && context.ConnectedDevice != null)
                {
                    context.DetailLabel.Text = $"{context.ConnectedDevice.DisplayName}\nCapture a frame to refresh the preview.";
                }
                else if (hasSelection && context.Selector.SelectedItem is CameraDeviceInfo pending)
                {
                    context.DetailLabel.Text = $"Selected: {pending.DisplayName}\nClick Connect Camera to assign it as the {context.RoleName.ToLowerInvariant()} view.";
                }
                else
                {
                    context.DetailLabel.Text = $"Choose which device should act as the {context.RoleName.ToLowerInvariant()} camera.";
                }
            }

            if (context.StatusLabel != null)
            {
                if (context.IsConnected)
                {
                    UpdateStatusLabel(context.StatusLabel,
                        statusPassBackground,
                        statusPassForeground,
                        $"{context.RoleName} connected ({context.ConnectedDevice.DisplayName}).");
                }
                else
                {
                    UpdateStatusLabel(context.StatusLabel,
                        statusNeutralBackground,
                        statusNeutralForeground,
                        $"{context.RoleName} camera not connected.");
                }
            }
        }

        private CameraDeviceInfo GetSelectedDevice(CameraContext context)
        {
            return context?.Selector?.SelectedItem as CameraDeviceInfo;
        }

        private CameraContext GetOppositeContext(CameraContext context)
        {
            if (context == null)
            {
                return null;
            }

            return ReferenceEquals(context, topCameraContext) ? frontCameraContext : topCameraContext;
        }

        private void ToggleCameraConnection(CameraContext context)
        {
            if (context == null)
            {
                return;
            }

            if (context.IsConnected)
            {
                outToLog($"[Camera] {context.RoleName} disconnected.", LogStatus.Info);
                context.Disconnect();
                ClearWorkflowPreview(context.Role);
                RefreshCameraUiState(context);
                UpdateInitSummary();
                return;
            }

            CameraDeviceInfo device = GetSelectedDevice(context);
            if (device == null)
            {
                MessageBox.Show(this,
                    $"Select a device for the {context.RoleName.ToLowerInvariant()} camera before connecting.",
                    "Cameras",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            CameraContext other = GetOppositeContext(context);
            if (other != null &&
                other.IsConnected &&
                string.Equals(other.ConnectedDevice.SerialNumber, device.SerialNumber, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this,
                    $"The selected device is already assigned to the {other.RoleName.ToLowerInvariant()} camera.",
                    "Cameras",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                MyCamera camera = new MyCamera();
                int createResult = camera.IMV_CreateHandle(IMVDefine.IMV_ECreateHandleMode.modeByIndex, device.Index);
                if (createResult != IMVDefine.IMV_OK)
                {
                    throw new InvalidOperationException($"CreateHandle failed ({createResult}).");
                }

                try
                {
                    int openResult = camera.IMV_Open();
                    if (openResult != IMVDefine.IMV_OK)
                    {
                        throw new InvalidOperationException($"Open camera failed ({openResult}).");
                    }
                }
                catch
                {
                    camera.IMV_DestroyHandle();
                    throw;
                }

                context.SetConnectedDevice(device, camera);
                RefreshCameraUiState(context);
                outToLog($"[Camera] {context.RoleName} connected to {device.DisplayName}.", LogStatus.Success);
            }
            catch (Exception ex)
            {
                context.Disconnect();
                ClearWorkflowPreview(context.Role);
                RefreshCameraUiState(context);
                MessageBox.Show(this, $"Failed to connect to camera: {ex.Message}", "Cameras",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                outToLog($"[Camera] {context.RoleName} connection failed: {ex.Message}", LogStatus.Error);
            }
            finally
            {
                UpdateInitSummary();
            }
        }

        private Bitmap CaptureCameraFrame(CameraContext context, int timeoutMs = 1500)
        {
            if (context == null || context.Camera == null || !context.IsConnected)
            {
                throw new InvalidOperationException("Camera is not connected.");
            }

            lock (context)
            {
                bool started = false;
                IMVDefine.IMV_Frame frame = new IMVDefine.IMV_Frame();
                try
                {
                    if (!context.Camera.IMV_IsGrabbing())
                    {
                        int startRes = context.Camera.IMV_StartGrabbing();
                        if (startRes != IMVDefine.IMV_OK)
                        {
                            throw new InvalidOperationException($"Start grabbing failed ({startRes}).");
                        }
                        started = true;
                    }

                    int frameRes = context.Camera.IMV_GetFrame(ref frame, (uint)timeoutMs);
                    if (frameRes != IMVDefine.IMV_OK)
                    {
                        throw new InvalidOperationException("Timeout waiting for frame.");
                    }

                    Bitmap bitmap = ConvertFrameToBitmap(context, frame);
                    if (bitmap == null)
                    {
                        throw new InvalidOperationException("Failed to convert captured frame to bitmap.");
                    }

                    return bitmap;
                }
                finally
                {
                    context.Camera.IMV_ReleaseFrame(ref frame);
                    if (started)
                    {
                        try
                        {
                            context.Camera.IMV_StopGrabbing();
                        }
                        catch
                        {
                            // ignore teardown errors
                        }
                    }
                }
            }
        }

        private void CaptureCameraPreview(CameraContext context)
        {
            if (context == null || !context.IsConnected)
            {
                return;
            }

            bool indicatorShown = false;
            Bitmap bitmap = null;
            try
            {
                ShowLoadingIndicator($"Capturing {context.RoleName.ToLowerInvariant()} preview...");
                indicatorShown = true;
                UseWaitCursor = true;

                bitmap = CaptureCameraFrame(context);
                if (bitmap != null)
                {
                    UpdateWorkflowPreview(context, bitmap);
                    if (context.DetailLabel != null && context.ConnectedDevice != null)
                    {
                        context.DetailLabel.Text = $"{context.ConnectedDevice.DisplayName}\nPreview captured at {DateTime.Now:T}.";
                    }

                    outToLog($"[Camera] Captured preview from {context.RoleName.ToLowerInvariant()} camera.", LogStatus.Success);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to capture preview: {ex.Message}", "Cameras",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                outToLog($"[Camera] Capture failed on {context.RoleName}: {ex.Message}", LogStatus.Error);
            }
            finally
            {
                bitmap?.Dispose();
                UseWaitCursor = false;
                if (indicatorShown)
                {
                    HideLoadingIndicator();
                }
            }
        }

        private DetectImg CaptureTopCameraImageForDetection()
        {
            Bitmap frame = null;
            try
            {
                frame = CaptureCameraFrame(topCameraContext, 2000);
                if (frame == null)
                {
                    throw new InvalidOperationException("Top camera returned an empty frame.");
                }

                string captureDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Captured");
                Directory.CreateDirectory(captureDirectory);
                string fileName = Path.Combine(captureDirectory, $"Top_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");

                Bitmap previewBitmap = (Bitmap)frame.Clone();
                ReplacePictureBoxImage(PB_OriginalImage, previewBitmap);

                frame.Save(fileName, ImageFormat.Png);
                Mat mat = BitmapToMat(frame);

                DetectImg detectImg = new DetectImg
                {
                    OriImg = mat,
                    OriImgPath = fileName,
                    ObjList = new List<ObjectInfo>(),
                    HasResults = false,
                    AttachmentPoints = new List<AttachmentPointInfo>(),
                    AttachmentCenter = new PointF(mat.Width / 2f, mat.Height / 2f),
                    FrontInspections = new List<FrontInspectionResult>(),
                    FrontInspectionComplete = false
                };

                outToLog($"[Camera] Captured top image for detection: {fileName}", LogStatus.Success);
                return detectImg;
            }
            finally
            {
                frame?.Dispose();
            }
        }

        private static Mat BitmapToMat(Bitmap bitmap)
        {
            if (bitmap == null)
            {
                throw new ArgumentNullException(nameof(bitmap));
            }

            using (Image<Bgr, byte> image = bitmap.ToImage<Bgr, byte>())
            {
                return image.Mat.Clone();
            }
        }

        private AttachmentOverlayResult BuildAttachmentOverlay(DetectImg image)
        {
            if (image?.OriImg == null)
            {
                return new AttachmentOverlayResult(null, new List<AttachmentPointInfo>(), PointF.Empty);
            }

            Size size = new Size(image.OriImg.Width, image.OriImg.Height);
            var (points, center) = ComputeAttachmentPoints(image.ObjList, size);

            if (image.AttachmentPoints != null && image.AttachmentPoints.Count > 0)
            {
                foreach (AttachmentPointInfo point in points)
                {
                    AttachmentPointInfo existing = image.AttachmentPoints.FirstOrDefault(p =>
                        ReferenceEquals(p.Source, point.Source) || p.Sequence == point.Sequence);
                    if (existing != null)
                    {
                        point.CapturedImagePath = existing.CapturedImagePath;
                        point.FrontInspection = existing.FrontInspection;
                    }
                }
            }

            using (Mat overlay = image.OriImg.Clone())
            {
                MCvScalar textColor = new MCvScalar(255, 255, 255);
                MCvScalar textOutlineColor = new MCvScalar(0, 0, 0);
                int radius = Math.Max(12, Math.Min(size.Width, size.Height) / 40);
                double fontScale = Math.Max(0.5, Math.Min(size.Width, size.Height) / 1200.0);
                int thickness = 2;

                // highlight reference center
                Point centerPoint = Point.Round(center);
                CvInvoke.DrawMarker(overlay, centerPoint, new MCvScalar(0, 165, 255), MarkerTypes.Star, 40, 2);

                foreach (AttachmentPointInfo point in points)
                {
                    Point circleCenter = Point.Round(point.Center);
                    Color statusColor = GetAttachmentStatusColor(point);
                    MCvScalar fillColor = ToScalar(LightenColor(statusColor, 0.35f));
                    MCvScalar outlineColor = ToScalar(point.Sequence == selectedFrontSequence ? Color.FromArgb(33, 150, 243) : Color.FromArgb(45, 45, 45));
                    int outlineThickness = point.Sequence == selectedFrontSequence ? 4 : 2;

                    CvInvoke.Circle(overlay, circleCenter, radius, fillColor, -1);
                    CvInvoke.Circle(overlay, circleCenter, radius, outlineColor, outlineThickness);

                    string label = $"{point.Sequence:00} | {point.AngleDegrees:+0.0;-0.0;+0.0}°";
                    int baseline = 0;
                    Size textSize = CvInvoke.GetTextSize(label, FontFace.HersheySimplex, fontScale, thickness, ref baseline);
                    int textX = (int)Math.Round(circleCenter.X - textSize.Width / 2.0);
                    int textY = (int)Math.Round((double)circleCenter.Y - radius - 10.0);
                    textX = Math.Max(0, Math.Min(size.Width - textSize.Width, textX));

                    CvInvoke.PutText(overlay, label, new Point(textX, textY), FontFace.HersheySimplex, fontScale, textOutlineColor, thickness + 1);
                    CvInvoke.PutText(overlay, label, new Point(textX, textY), FontFace.HersheySimplex, fontScale, textColor, thickness);
                }

                Bitmap overlayBitmap = overlay.ToBitmap();
                return new AttachmentOverlayResult(overlayBitmap, points, center);
            }
        }

        private (List<AttachmentPointInfo> Points, PointF Center) ComputeAttachmentPoints(List<ObjectInfo> objects, Size imageSize)
        {
            List<AttachmentPointInfo> result = new List<AttachmentPointInfo>();
            if (objects == null || objects.Count == 0)
            {
                return (result, new PointF(imageSize.Width / 2f, imageSize.Height / 2f));
            }

            List<(ObjectInfo Obj, PointF Center)> rawPoints = new List<(ObjectInfo, PointF)>();
            foreach (ObjectInfo obj in objects)
            {
                Rectangle rect = obj.DisplayRec;
                if (rect.Width <= 0 || rect.Height <= 0)
                {
                    continue;
                }

                PointF center = new PointF(rect.Left + rect.Width / 2f, rect.Top + rect.Height / 2f);
                rawPoints.Add((obj, center));
            }

            if (rawPoints.Count == 0)
            {
                return (result, new PointF(imageSize.Width / 2f, imageSize.Height / 2f));
            }

            PointF referenceCenter = new PointF(imageSize.Width / 2f, imageSize.Height / 2f);

            List<double> angles = new List<double>(rawPoints.Count);
            foreach ((ObjectInfo obj, PointF center) in rawPoints)
            {
                double angle = CalculateAttachmentAngle(center, referenceCenter);
                obj.DisplayAngle = (float)angle;
                angles.Add(angle);
            }

            var ordered = angles
                .Select((angle, index) => new { angle, index })
                .OrderByDescending(item => item.angle)
                .ToList();

            int sequence = 1;
            foreach (var entry in ordered)
            {
                (ObjectInfo obj, PointF center) = rawPoints[entry.index];
                result.Add(new AttachmentPointInfo(obj, center, entry.angle, sequence));
                sequence++;
            }

            return (result, referenceCenter);
        }

        private static double CalculateAttachmentAngle(PointF point, PointF center)
        {
            double dx = point.X - center.X;
            double dy = point.Y - center.Y;
            if (Math.Abs(dx) < double.Epsilon && Math.Abs(dy) < double.Epsilon)
            {
                return 0.0;
            }

            double angleRad = Math.Atan2(dx, -dy);
            double angleDeg = angleRad * 180.0 / Math.PI;
            return NormalizeSignedAngle(angleDeg);
        }

        private Color GetAttachmentStatusColor(AttachmentPointInfo point)
        {
            if (point?.FrontInspection == null)
            {
                return Color.FromArgb(96, 125, 139);
            }

            return point.FrontInspection.HasDefects
                ? Color.FromArgb(229, 57, 53)
                : Color.FromArgb(46, 204, 113);
        }

        private static MCvScalar ToScalar(Color color)
        {
            return new MCvScalar(color.B, color.G, color.R);
        }

        private static Color LightenColor(Color color, float amount)
        {
            amount = Math.Max(0f, Math.Min(1f, amount));
            int r = (int)(color.R + (255 - color.R) * amount);
            int g = (int)(color.G + (255 - color.G) * amount);
            int b = (int)(color.B + (255 - color.B) * amount);
            return Color.FromArgb(r, g, b);
        }

        private static double NormalizeSignedAngle(double angle)
        {
            while (angle > 180.0)
            {
                angle -= 360.0;
            }

            while (angle <= -180.0)
            {
                angle += 360.0;
            }

            return angle;
        }

        private unsafe Bitmap ConvertFrameToBitmap(CameraContext context, IMVDefine.IMV_Frame frame)
        {
            int width = (int)frame.frameInfo.width;
            int height = (int)frame.frameInfo.height;

            IntPtr sourcePtr = frame.pData;
            int bytesPerPixel;
            bool usedConversion = false;

            if (frame.frameInfo.pixelFormat == IMVDefine.IMV_EPixelType.gvspPixelMono8)
            {
                bytesPerPixel = 1;
            }
            else if (frame.frameInfo.pixelFormat == IMVDefine.IMV_EPixelType.gvspPixelBGR8)
            {
                bytesPerPixel = 3;
            }
            else
            {
                int requiredSize = width * height * 3;
                if (context.ConversionBuffer == IntPtr.Zero || requiredSize > context.ConversionBufferSize)
                {
                    if (context.ConversionBuffer != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(context.ConversionBuffer);
                    }

                    context.ConversionBuffer = Marshal.AllocHGlobal(requiredSize);
                    context.ConversionBufferSize = requiredSize;
                }

                IMVDefine.IMV_PixelConvertParam convertParam = new IMVDefine.IMV_PixelConvertParam
                {
                    nWidth = frame.frameInfo.width,
                    nHeight = frame.frameInfo.height,
                    ePixelFormat = frame.frameInfo.pixelFormat,
                    pSrcData = frame.pData,
                    nSrcDataLen = frame.frameInfo.size,
                    nPaddingX = frame.frameInfo.paddingX,
                    nPaddingY = frame.frameInfo.paddingY,
                    eBayerDemosaic = IMVDefine.IMV_EBayerDemosaic.demosaicNearestNeighbor,
                    eDstPixelFormat = IMVDefine.IMV_EPixelType.gvspPixelBGR8,
                    pDstBuf = context.ConversionBuffer,
                    nDstBufSize = (uint)context.ConversionBufferSize
                };

                int convertResult = context.Camera.IMV_PixelConvert(ref convertParam);
                if (convertResult != IMVDefine.IMV_OK)
                {
                    throw new InvalidOperationException($"Pixel convert failed ({convertResult}).");
                }

                sourcePtr = context.ConversionBuffer;
                bytesPerPixel = 3;
                usedConversion = true;
            }

            PixelFormat bitmapFormat = bytesPerPixel == 1 ? PixelFormat.Format8bppIndexed : PixelFormat.Format24bppRgb;
            Bitmap bitmap = new Bitmap(width, height, bitmapFormat);
            Rectangle rect = new Rectangle(0, 0, width, height);
            BitmapData bmpData = bitmap.LockBits(rect, ImageLockMode.WriteOnly, bitmapFormat);

            int srcStride = usedConversion
                ? width * bytesPerPixel
                : (int)(width * bytesPerPixel + frame.frameInfo.paddingX);
            int destStride = bmpData.Stride;
            int rowLength = width * bytesPerPixel;

            byte* srcBase = (byte*)sourcePtr.ToPointer();
            byte* destBase = (byte*)bmpData.Scan0.ToPointer();

            for (int y = 0; y < height; y++)
            {
                byte* srcRow = srcBase + y * srcStride;
                byte* destRow = destBase + y * destStride;
                Buffer.MemoryCopy(srcRow, destRow, destStride, rowLength);
            }

            bitmap.UnlockBits(bmpData);

            if (bitmapFormat == PixelFormat.Format8bppIndexed)
            {
                ColorPalette palette = bitmap.Palette;
                for (int i = 0; i < 256; i++)
                {
                    palette.Entries[i] = Color.FromArgb(i, i, i);
                }
                bitmap.Palette = palette;
            }

            return bitmap;
        }

        private void UpdateWorkflowPreview(CameraContext context, Bitmap source)
        {
            if (context == null || source == null)
            {
                return;
            }

            Rectangle area = new Rectangle(0, 0, source.Width, source.Height);
            Bitmap clone = source.Clone(area, source.PixelFormat);

            if (context.Role == CameraRole.Top)
            {
                ReplacePictureBoxImage(PB_OriginalImage, clone);
            }
            else
            {
                ReplacePictureBoxImage(PB_FrontPreview, clone);
            }
        }

        private void ClearWorkflowPreview(CameraRole role)
        {
            if (role == CameraRole.Top)
            {
                ReplacePictureBoxImage(PB_OriginalImage, null);
            }
            else
            {
                ReplacePictureBoxImage(PB_FrontPreview, null);
            }
        }

        private static void ReplacePictureBoxImage(PictureBox pictureBox, Image newImage)
        {
            if (pictureBox == null)
            {
                newImage?.Dispose();
                return;
            }

            Image previous = pictureBox.Image;
            pictureBox.Image = newImage;

            if (previous != null && !ReferenceEquals(previous, newImage))
            {
                previous.Dispose();
            }
        }

        private static string DescribeInterface(IMVDefine.IMV_EInterfaceType interfaceType)
        {
            switch (interfaceType)
            {
                case IMVDefine.IMV_EInterfaceType.interfaceTypeGige:
                    return "GigE";
                case IMVDefine.IMV_EInterfaceType.interfaceTypeUsb3:
                    return "USB3";
                case IMVDefine.IMV_EInterfaceType.interfaceTypeCL:
                    return "CameraLink";
                case IMVDefine.IMV_EInterfaceType.interfaceTypePCIe:
                    return "PCIe";
                case IMVDefine.IMV_EInterfaceType.interfaceTypeAll:
#if DEBUG
                    return "All";
#else
                    return "Unknown";
#endif
                default:
                    return interfaceType.ToString();
            }
        }

        private void UpdateModelStatus(ModelContext context, string message, Color background, Color foreground)
        {
            if (context?.StatusDisplay == null)
            {
                return;
            }

            context.StatusDisplay.BackColor = background;
            context.StatusDisplay.ForeColor = foreground;
            context.StatusDisplay.Text = message;
        }

        private void UpdateInitSummary()
        {
            if (LBL_InitSummary == null)
            {
                return;
            }
            bool attachmentLoaded = attachmentContext?.IsLoaded == true;
            bool defectLoaded = defectContext?.IsLoaded == true;
            string attachmentLabel = attachmentLoaded ? Path.GetFileName(attachmentContext.LoadedPath) : "Attachment not loaded";
            string defectLabel = defectLoaded ? Path.GetFileName(defectContext.LoadedPath) : "Defect not loaded";
            bool topConnected = topCameraContext?.IsConnected == true;
            bool frontConnected = frontCameraContext?.IsConnected == true;
            string cameraLabel;
            if (topConnected && frontConnected)
            {
                string topName = topCameraContext?.ConnectedDevice?.DisplayName ?? "Top camera";
                string frontName = frontCameraContext?.ConnectedDevice?.DisplayName ?? "Front camera";
                cameraLabel = $"Cameras ready ({topName}; {frontName})";
            }
            else if (!topConnected && !frontConnected)
            {
                cameraLabel = "Cameras disconnected";
            }
            else if (!topConnected)
            {
                cameraLabel = "Top camera disconnected";
            }
            else
            {
                cameraLabel = "Front camera disconnected";
            }
            bool turntableConnected = turntableController?.IsConnected == true;
            bool turntableHomed = turntableController?.IsHomed == true;
            double? offset = turntableController?.LastOffsetAngle;
            string turntableLabel;
            if (!turntableConnected)
            {
                turntableLabel = "Turntable disconnected";
            }
            else if (turntableHomed)
            {
                turntableLabel = offset.HasValue
                    ? $"Turntable homed ({offset.Value:0.00} deg)"
                    : "Turntable homed";
            }
            else
            {
                turntableLabel = "Turntable connected (not homed)";
            }
            bool ready = attachmentLoaded && defectLoaded && topConnected && frontConnected && turntableConnected && turntableHomed;
            LBL_InitSummary.Text = ready
                ? $"Ready: {attachmentLabel} | {defectLabel} | {cameraLabel} | {turntableLabel}"
                : $"Attachment: {attachmentLabel} | Defect: {defectLabel} | {cameraLabel} | {turntableLabel}";
            LBL_InitSummary.BackColor = ready ? statusPassBackground : statusNeutralBackground;
            LBL_InitSummary.ForeColor = ready ? statusPassForeground : statusNeutralForeground;
        }

            private void PopulateTurntablePorts()

        {

            if (CB_TurntablePort == null)

            {

                return;

            }

            string previouslySelected = CB_TurntablePort.SelectedItem as string;

            string[] ports = SerialPort.GetPortNames()

                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)

                .ToArray();

            if (turntableController != null &&

                turntableController.IsConnected &&

                !string.IsNullOrEmpty(turntableController.PortName) &&

                !ports.Contains(turntableController.PortName, StringComparer.OrdinalIgnoreCase))

            {

                ports = ports.Concat(new[] { turntableController.PortName }).ToArray();

            }

            CB_TurntablePort.BeginUpdate();

            CB_TurntablePort.Items.Clear();

            foreach (string port in ports)

            {

                CB_TurntablePort.Items.Add(port);

            }

            if (!string.IsNullOrWhiteSpace(previouslySelected) &&

                ports.Contains(previouslySelected, StringComparer.OrdinalIgnoreCase))

            {

                CB_TurntablePort.SelectedItem = previouslySelected;

            }

            else if (turntableController?.IsConnected == true && !string.IsNullOrEmpty(turntableController.PortName))

            {

                CB_TurntablePort.SelectedItem = turntableController.PortName;

            }

            else if (ports.Length > 0)

            {

                CB_TurntablePort.SelectedIndex = 0;

            }

            else

            {

                CB_TurntablePort.SelectedIndex = -1;

            }

            CB_TurntablePort.EndUpdate();

            UpdateTurntableConnectionUI(turntableController?.IsConnected == true, turntableController?.PortName);

        }



        private void UpdateTurntableConnectionUI(bool connected, string portName = null)

        {

            if (CB_TurntablePort == null)

            {

                return;

            }

            CB_TurntablePort.Enabled = !connected;

            BT_TurntableRefresh.Enabled = !connected;

            bool hasPortSelection = CB_TurntablePort != null && CB_TurntablePort.Items.Count > 0;

            BT_TurntableConnect.Enabled = connected || hasPortSelection;

            BT_TurntableConnect.Text = connected ? "Disconnect" : "Connect";

            BT_TurntableHome.Enabled = connected;

            if (connected)

            {

                string displayPort = portName ?? turntableController?.PortName ?? "Unknown";

                UpdateTurntableStatus($"Connected ({displayPort})", statusPassBackground, statusPassForeground);

            }

            else

            {

                UpdateTurntableStatus("Disconnected.", statusNeutralBackground, statusNeutralForeground);

            }

            UpdateInitSummary();

        }



        private void UpdateTurntableStatus(string message, Color background, Color foreground)

        {

            if (LBL_TurntableStatus == null)

            {

                return;

            }

            if (InvokeRequired)

            {

                BeginInvoke(new Action(() => UpdateTurntableStatus(message, background, foreground)));

                return;

            }

            UpdateStatusLabel(LBL_TurntableStatus, background, foreground, message);

        }



        private void TurntableController_MessageReceived(string message)

        {

            if (string.IsNullOrWhiteSpace(message))

            {

                return;

            }

            string normalized = TurntableController.NormalizeMessage(message);

            if (string.IsNullOrEmpty(normalized))

            {

                return;

            }

            void Process()

            {

                outToLog($"[Turntable] {normalized}", LogStatus.Info);

                if (normalized.StartsWith("CR+ERR", StringComparison.OrdinalIgnoreCase))

                {

                    UpdateTurntableStatus(normalized, statusFailBackground, statusFailForeground);

                }

                else if (normalized.IndexOf("OffsetAngle", StringComparison.OrdinalIgnoreCase) >= 0)

                {

                    if (turntableController?.LastOffsetAngle != null)

                    {

                        UpdateTurntableStatus(

                            $"Homed ({turntableController.LastOffsetAngle.Value:0.00} deg)",

                            statusPassBackground,

                            statusPassForeground);

                    }

                    UpdateInitSummary();

                }

            }

            if (InvokeRequired)

            {

                BeginInvoke((Action)Process);

            }

            else

            {

                Process();

            }

        }



        private void BT_TurntableRefresh_Click(object sender, EventArgs e)

        {

            PopulateTurntablePorts();

        }



        private void BT_TurntableConnect_Click(object sender, EventArgs e)

        {

            if (turntableController == null)

            {

                return;

            }

            if (turntableController.IsConnected)

            {

                turntableController.Disconnect();

                PopulateTurntablePorts();

                outToLog("[Turntable] Disconnected.", LogStatus.Info);

                return;

            }

            string port = CB_TurntablePort?.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(port))

            {

                MessageBox.Show(this, "Select a COM port before connecting.", "Turntable", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return;

            }

            try

            {

                turntableController.Connect(port);

                UpdateTurntableConnectionUI(true, port);

                outToLog($"[Turntable] Connected to {port}.", LogStatus.Success);

            }

            catch (Exception ex)

            {

                UpdateTurntableConnectionUI(false);

                outToLog($"[Turntable] Connection failed: {ex.Message}", LogStatus.Error);

                MessageBox.Show(this, ex.Message, "Turntable Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }



        private async void BT_TurntableHome_Click(object sender, EventArgs e)

        {

            await RunTurntableHomeAsync();

        }



        private async Task RunTurntableHomeAsync()

        {

            if (turntableController == null || !turntableController.IsConnected)

            {

                MessageBox.Show(this, "Connect to the turntable before homing.", "Turntable", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                FocusInitializeTab();

                return;

            }

            BT_TurntableHome.Enabled = false;

            bool indicatorShown = false;

            try

            {

                ShowLoadingIndicator("Homing turntable...");

                indicatorShown = true;

                using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(20)))

                {

                    TurntableHomeResult result = await turntableController.HomeAsync(cts.Token);

                    if (result.Success)

                    {

                        UpdateTurntableStatus(

                            result.OffsetDegrees.HasValue

                                ? $"Homed ({result.OffsetDegrees.Value:0.00} deg)"

                                : "Homed.",

                            statusPassBackground,

                            statusPassForeground);

                        outToLog(

                            result.OffsetDegrees.HasValue

                                ? $"[Turntable] Homing complete. Offset {result.OffsetDegrees.Value:0.00} deg."

                                : "[Turntable] Homing complete.",

                            LogStatus.Success);

                    }

                    else

                    {

                        UpdateTurntableStatus(result.Message, statusFailBackground, statusFailForeground);

                        outToLog($"[Turntable] Homing failed: {result.Message}", LogStatus.Error);

                        FocusInitializeTab();

                    }

                }

            }

            catch (OperationCanceledException)

            {

                UpdateTurntableStatus("Homing timed out.", statusFailBackground, statusFailForeground);

                outToLog("[Turntable] Homing timed out.", LogStatus.Error);

                FocusInitializeTab();

            }

            catch (Exception ex)

            {

                UpdateTurntableStatus($"Homing error: {ex.Message}", statusFailBackground, statusFailForeground);

                outToLog($"[Turntable] Homing error: {ex.Message}", LogStatus.Error);

                FocusInitializeTab();

            }

            finally

            {

                if (indicatorShown)

                {

                    HideLoadingIndicator();

                }

                BT_TurntableHome.Enabled = turntableController?.IsConnected == true;

                UpdateInitSummary();

            }

        }



        private void FocusInitializeTab()

        {

            if (leftTabs != null && tabInitialize != null && leftTabs.TabPages.Contains(tabInitialize))

            {

                leftTabs.SelectedTab = tabInitialize;

            }

        }



        private void ShowLoadingIndicator(string message)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                Invoke(new Action<string>(ShowLoadingIndicator), message);
                return;
            }

            loadingRequestCount++;
            loadingSpinner.Visible = true;
            loadingIndicatorPanel.Visible = true;
            if (toolTip != null)
            {
                toolTip.SetToolTip(loadingIndicatorPanel, message);
            }
            PositionLoadingIndicator();
            loadingIndicatorPanel.BringToFront();
            UseWaitCursor = true;
        }

        private void HideLoadingIndicator()
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                Invoke(new Action(HideLoadingIndicator));
                return;
            }

            loadingRequestCount = Math.Max(0, loadingRequestCount - 1);
            if (loadingRequestCount == 0)
            {
                loadingIndicatorPanel.Visible = false;
                loadingSpinner.Visible = false;
                UseWaitCursor = false;
            }
        }

        private void BT_Detect_Click(object sender, EventArgs e)
        {
            if (!EnsureReadyForDetection())
            {
                return;
            }

            CancelAttachmentSequence();

            DetectImg capturedImage;
            try
            {
                capturedImage = CaptureTopCameraImageForDetection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to capture top camera image: {ex.Message}", "Detection",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                outToLog($"[Camera] Top capture failed: {ex.Message}", LogStatus.Error);
                return;
            }

            currentImage?.Dispose();
            currentImage = capturedImage;
            DisplayOriginalImage();
            ClearDetectionVisuals();

            outToLog("Running detection...", LogStatus.Progress);
            SetLogicStatusNeutral("Running detection...");
            ShowLoadingIndicator("Running detection...");

            BT_Detect.Enabled = false;
            Task.Run(() => DetectThreadImpl());
        }

        private bool EnsureReadyForDetection()
        {
            if (SolDL == null)
            {
                outToLog("SDK not initialized.", LogStatus.Error);
                return false;
            }

            if (attachmentContext?.IsLoaded != true)
            {
                outToLog("Please load the attachment project before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (defectContext?.IsLoaded != true)
            {
                outToLog("Please load the defect project before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (topCameraContext?.IsConnected != true)
            {
                outToLog("Connect the top camera before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (frontCameraContext?.IsConnected != true)
            {
                outToLog("Connect the front camera before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (turntableController?.IsConnected != true)
            {
                outToLog("Connect to the turntable before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (!turntableController.IsHomed)
            {
                outToLog("Home the turntable before running detection.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            if (string.IsNullOrWhiteSpace(loadedProjectPath))
            {
                outToLog("Attachment project path is missing.", LogStatus.Warning);
                FocusInitializeTab();
                return false;
            }

            return true;
        }



        private void DetectThreadImpl()
        {
            var targetImage = currentImage;
            if (targetImage?.OriImg == null)
            {
                CancelAttachmentSequence();
                BeginInvoke(new MethodInvoker(() =>
                {
                    outToLog("No image available for detection.", LogStatus.Warning);
                    BT_Detect.Enabled = true;
                    SetLogicStatusNeutral("Awaiting defect inspection.");
                    HideLoadingIndicator();
                }));
                return;
            }

            try
            {
                Stopwatch countTime = Stopwatch.StartNew();
                List<ObjectInfo> detections;
                using (Mat matImg = targetImage.OriImg.Clone())
                {
                    SolDL.Detect(matImg, out detections);
                }
                countTime.Stop();

                List<ObjectInfo> normalizedDetections = detections ?? new List<ObjectInfo>();
                AttachmentOverlayResult overlayResult = null;

                if (targetImage == currentImage)
                {
                    targetImage.ObjList = normalizedDetections;
                    overlayResult = BuildAttachmentOverlay(targetImage);
                }

                BeginInvoke(new MethodInvoker(() =>
                {
                    if (targetImage == currentImage)
                    {
                        currentImage.HasResults = true;
                        if (overlayResult != null)
                        {
                            currentImage.AttachmentPoints = overlayResult.Points;
                            currentImage.AttachmentCenter = overlayResult.Center;
                            if (overlayResult.OverlayBitmap != null)
                            {
                                ReplacePictureBoxImage(PB_OriginalImage, overlayResult.OverlayBitmap);
                            }
                            else
                            {
                                ReplacePictureBoxImage(PB_OriginalImage, currentImage.OriImg.ToBitmap());
                            }
                        }
                        else
                        {
                            currentImage.AttachmentPoints = new List<AttachmentPointInfo>();
                            ReplacePictureBoxImage(PB_OriginalImage, currentImage.OriImg.ToBitmap());
                        }

                        PopulateDetectionGrid(currentImage.ObjList);
                        if (CB_Field.SelectedValue is LogicField selectedField && selectedField == LogicField.ClassName)
                        {
                            UpdateValueSuggestions(LogicField.ClassName, CB_Value.Text);
                        }
                        UpdateLogicEvaluation();
                        outToLog($"Detection finished in {countTime.ElapsedMilliseconds} ms for {Path.GetFileName(currentImage.OriImgPath)}.", LogStatus.Success);
                    }
                    StartAttachmentCaptureSequence(currentImage);
                    BT_Detect.Enabled = true;
                    HideLoadingIndicator();
                }));
            }
            catch (Exception ex)
            {
                BeginInvoke(new MethodInvoker(() =>
                {
                    outToLog($"Detection failed: {ex.Message}", LogStatus.Error);
                    BT_Detect.Enabled = true;
                    if (currentImage != null)
                    {
                        currentImage.HasResults = false;
                    }
                    SetLogicStatusNeutral("Detection failed.");
                    HideLoadingIndicator();
                }));
            }
        }

        protected void UpdateClassNames(List<string> classNames)
        {
            classNameList = classNames;
            if (CB_Field != null && CB_Field.Visible && CB_Field.SelectedValue is LogicField field && field == LogicField.ClassName)
            {
                UpdateValueSuggestions(LogicField.ClassName, CB_Value?.Text);
            }
        }

        public List<string> GetClassNames()
        {
            SolVision.TaskProcess target = defectContext?.Process ?? SolDL;
            if (target != null)
            {
                try
                {
                    List<string> classNames = target.GetClassNames() ?? new List<string>();
                    classNames.RemoveAll(p => p.Equals("BackGround", StringComparison.OrdinalIgnoreCase) || p.Equals("Unknown", StringComparison.OrdinalIgnoreCase));
                    if (classNames.Count == 0)
                    {
                        return new List<string> { "Object" };
                    }
                    return classNames;
                }
                catch
                {
                    return new List<string> { "Object" };
                }
            }

            return new List<string> { "Object" };
        }

        private void RefreshDefectClassNamesFromModel()
        {
            classNameList = GetClassNames() ?? new List<string>();
            if (CB_Field != null && CB_Field.SelectedValue is LogicField field && field == LogicField.ClassName)
            {
                UpdateValueSuggestions(LogicField.ClassName, CB_Value?.Text);
            }
        }

        private void DG_Detections_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || currentImage?.ObjList == null || currentImage.ObjList.Count == 0)
            {
                return;
            }

            if (e.RowIndex >= currentImage.ObjList.Count)
            {
                return;
            }

            ObjectInfo selected = currentImage.ObjList[e.RowIndex];
            RenderSingleDetection(selected);
        }

        private void RenderAllDetections()
        {
            if (currentImage?.OriImg == null)
            {
                ReplacePictureBoxImage(PB_OriginalImage, null);
                return;
            }

            AttachmentOverlayResult overlayResult = BuildAttachmentOverlay(currentImage);
            currentImage.AttachmentPoints = overlayResult.Points;
            currentImage.AttachmentCenter = overlayResult.Center;
            if (overlayResult.OverlayBitmap != null)
            {
                ReplacePictureBoxImage(PB_OriginalImage, overlayResult.OverlayBitmap);
            }
            else
            {
                ReplacePictureBoxImage(PB_OriginalImage, currentImage.OriImg.ToBitmap());
            }
        }

        private void RenderSingleDetection(ObjectInfo target)
        {
            if (currentImage?.OriImg == null || target == null)
            {
                return;
            }

            AttachmentOverlayResult overlayResult = BuildAttachmentOverlay(currentImage);
            if (overlayResult.OverlayBitmap == null)
            {
                ReplacePictureBoxImage(PB_OriginalImage, currentImage.OriImg.ToBitmap());
                return;
            }

            Rectangle rect = target.DisplayRec;
            PointF center = new PointF(rect.Left + rect.Width / 2f, rect.Top + rect.Height / 2f);

            using (Graphics g = Graphics.FromImage(overlayResult.OverlayBitmap))
            using (Pen highlight = new Pen(Color.OrangeRed, 3))
            {
                g.DrawEllipse(highlight, center.X - 28, center.Y - 28, 56, 56);
            }

            currentImage.AttachmentPoints = overlayResult.Points;
            currentImage.AttachmentCenter = overlayResult.Center;
            ReplacePictureBoxImage(PB_OriginalImage, overlayResult.OverlayBitmap);
        }

        private void DisplayOriginalImage()
        {
            if (currentImage?.OriImg == null)
            {
                ReplacePictureBoxImage(PB_OriginalImage, null);
                return;
            }

            ReplacePictureBoxImage(PB_OriginalImage, currentImage.OriImg.ToBitmap());
        }

        private void PopulateDetectionGrid(List<ObjectInfo> detections)
        {
            DG_Detections.Rows.Clear();
            if (currentImage?.AttachmentPoints != null && currentImage.AttachmentPoints.Count > 0)
            {
                foreach (AttachmentPointInfo point in currentImage.AttachmentPoints)
                {
                    ObjectInfo obj = point.Source;
                    Rectangle rec = obj.DisplayRec;
                    string bounds = $"{rec.X},{rec.Y},{rec.Width},{rec.Height}";
                    string centerText = $"{point.Center.X:0},{point.Center.Y:0}";
                    string className = string.IsNullOrWhiteSpace(obj.name) ? "(unknown)" : obj.name;
                    float score = obj.confidence != 0 ? obj.confidence : obj.classifyScore;
                    DG_Detections.Rows.Add(point.Sequence, className, score.ToString("0.00"), point.AngleDegrees.ToString("+0.0;-0.0;+0.0"), centerText, bounds);
                }
                return;
            }

            if (detections == null || detections.Count == 0)
            {
                return;
            }

            PointF fallbackCenter = currentImage?.OriImg != null
                ? new PointF(currentImage.OriImg.Width / 2f, currentImage.OriImg.Height / 2f)
                : new PointF(0, 0);

            int sequence = 1;
            foreach (ObjectInfo obj in detections)
            {
                Rectangle rec = obj.DisplayRec;
                string bounds = $"{rec.X},{rec.Y},{rec.Width},{rec.Height}";
                PointF center = new PointF(rec.Left + rec.Width / 2f, rec.Top + rec.Height / 2f);
                string centerText = $"{center.X:0},{center.Y:0}";
                string className = string.IsNullOrWhiteSpace(obj.name) ? "(unknown)" : obj.name;
                float score = obj.confidence != 0 ? obj.confidence : obj.classifyScore;
                double angle = CalculateAttachmentAngle(center, fallbackCenter);
                DG_Detections.Rows.Add(sequence, className, score.ToString("0.00"), angle.ToString("+0.0;-0.0;+0.0"), centerText, bounds);
                sequence++;
            }
        }

        private void ClearDetectionVisuals()
        {
            if (currentImage != null)
            {
                currentImage.ObjList.Clear();
                currentImage.HasResults = false;
                currentImage.AttachmentPoints?.Clear();
                currentImage.AttachmentCenter = new PointF();
                if (currentImage.FrontInspections != null)
                {
                    foreach (FrontInspectionResult inspection in currentImage.FrontInspections)
                    {
                        inspection?.Dispose();
                    }
                    currentImage.FrontInspections.Clear();
                }
                currentImage.FrontInspectionComplete = false;
            }
            DG_Detections.Rows.Clear();
            selectedFrontSequence = -1;
            SetPictureBoxImage(PB_FrontPreview, null);
            RefreshFrontGallery();
            RefreshTopOverlay();
            UpdateFrontPreviewAndLedger();
            SetLogicStatusNeutral("Awaiting defect inspection.");
        }

        private void CancelAttachmentSequence()
        {
            CancellationTokenSource cts = Interlocked.Exchange(ref captureSequenceCts, null);
            if (cts != null)
            {
                cts.Cancel();
                cts.Dispose();
            }
        }

        private void StartAttachmentCaptureSequence(DetectImg image)
        {
            if (image == null)
            {
                return;
            }

            if (image.AttachmentPoints == null || image.AttachmentPoints.Count == 0)
            {
                image.FrontInspectionComplete = true;
                if (image.FrontInspections != null && image.FrontInspections.Count > 0)
                {
                    foreach (FrontInspectionResult inspection in image.FrontInspections)
                    {
                        inspection?.Dispose();
                    }
                    image.FrontInspections.Clear();
                }
                UpdateDefectClassNamesFromInspections(image.FrontInspections);
                UpdateLogicEvaluation();
                return;
            }

            if (image.FrontInspections != null && image.FrontInspections.Count > 0)
            {
                foreach (FrontInspectionResult inspection in image.FrontInspections)
                {
                    inspection?.Dispose();
                }
                image.FrontInspections.Clear();
            }
            image.FrontInspectionComplete = false;

            selectedFrontSequence = -1;
            RefreshFrontGallery();
            UpdateFrontPreviewAndLedger();

            if (turntableController == null || !turntableController.IsConnected)
            {
                outToLog("[Sequence] Turntable not connected; skipping rotation capture.", LogStatus.Warning);
                return;
            }

            if (!turntableController.IsHomed)
            {
                outToLog("[Sequence] Turntable not homed; skipping rotation capture.", LogStatus.Warning);
                return;
            }

            if (frontCameraContext == null || !frontCameraContext.IsConnected)
            {
                outToLog("[Sequence] Front camera not connected; skipping rotation capture.", LogStatus.Warning);
                return;
            }

            List<AttachmentPointInfo> snapshot = image.AttachmentPoints.ToList();
            if (snapshot.Count == 0)
            {
                return;
            }

            var cts = new CancellationTokenSource();
            CancellationTokenSource previous = Interlocked.Exchange(ref captureSequenceCts, cts);
            if (previous != null)
            {
                previous.Cancel();
                previous.Dispose();
            }

            outToLog($"[Sequence] Starting capture sequence for {snapshot.Count} attachment point(s).", LogStatus.Progress);

            Task.Run(async () =>
            {
                try
                {
                    await RunAttachmentSequenceAsync(image, snapshot, cts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog("[Sequence] Capture sequence cancelled.", LogStatus.Warning);
                        }));
                    }
                }
                catch (Exception ex)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Sequence] Capture sequence failed: {ex.Message}", LogStatus.Error);
                        }));
                    }
                }
                finally
                {
                    if (Interlocked.CompareExchange(ref captureSequenceCts, null, cts) == cts)
                    {
                        cts.Dispose();
                    }
                }
            });
        }

        private async Task RunAttachmentSequenceAsync(DetectImg sourceImage, List<AttachmentPointInfo> points, CancellationToken token)
        {
            if (sourceImage == null || points == null || points.Count == 0)
            {
                return;
            }

            TurntableController controller = turntableController;
            CameraContext frontContext = frontCameraContext;

            if (controller == null || frontContext == null || !frontContext.IsConnected)
            {
                return;
            }

            string captureDirectory = EnsureFrontCaptureDirectory(sourceImage);
            double currentAngle = 0.0;
            const double orientationMultiplier = -1.0; // front camera is inverted, rotate in the opposite direction first
            List<FrontCaptureTicket> captureTickets = new List<FrontCaptureTicket>();

            foreach (AttachmentPointInfo point in points)
            {
                token.ThrowIfCancellationRequested();

                if (!ReferenceEquals(sourceImage, currentImage))
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog("[Sequence] New detection detected; stopping previous capture sequence.", LogStatus.Warning);
                        }));
                    }
                    return;
                }

                double targetAngle = NormalizeSignedAngle(point.AngleDegrees);
                double physicalTarget = NormalizeSignedAngle(targetAngle * orientationMultiplier);
                double moveDelta = NormalizeSignedAngle(physicalTarget - currentAngle);

                if (Math.Abs(moveDelta) > 0.1)
                {
                    try
                    {
                        using (CancellationTokenSource moveCts = CancellationTokenSource.CreateLinkedTokenSource(token))
                        {
                            moveCts.CancelAfter(TimeSpan.FromSeconds(45));
                            string summary = await controller.MoveRelativeAsync(moveDelta, moveCts.Token).ConfigureAwait(false);
                            if (!IsDisposed)
                            {
                                BeginInvoke(new MethodInvoker(() =>
                                {
                                    outToLog($"[Turntable] Index {point.Sequence}: {summary}", LogStatus.Info);
                                }));
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        if (!IsDisposed)
                        {
                            BeginInvoke(new MethodInvoker(() =>
                            {
                                outToLog($"[Turntable] Failed to move to index {point.Sequence}: {ex.Message}", LogStatus.Error);
                            }));
                        }
                        return;
                    }
                }

                currentAngle = physicalTarget;

                token.ThrowIfCancellationRequested();

                try
                {
                    Bitmap frame = CaptureCameraFrame(frontContext, 2000);
                    string angleToken = BuildAngleToken(physicalTarget);
                    string filePath = Path.Combine(captureDirectory, $"Front_Index{point.Sequence:00}_{angleToken}.png");
                    frame.Save(filePath, ImageFormat.Png);
                    point.CapturedImagePath = filePath;
                    captureTickets.Add(new FrontCaptureTicket(point.Sequence, filePath, physicalTarget, DateTime.Now));

                    Bitmap uiBitmap = (Bitmap)frame.Clone();
                    frame.Dispose();

                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            UpdateWorkflowPreview(frontContext, uiBitmap);
                            outToLog($"[Capture] Saved front image for index {point.Sequence} at {physicalTarget:+0.0;-0.0;+0.0} deg -> {filePath}", LogStatus.Success);
                            uiBitmap.Dispose();
                        }));
                    }
                    else
                    {
                        uiBitmap.Dispose();
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Capture] Failed to capture front image for index {point.Sequence}: {ex.Message}", LogStatus.Error);
                        }));
                    }
                }
            }

            double returnDelta = NormalizeSignedAngle(-currentAngle);
            Task<List<FrontInspectionResult>> processTask = Task.FromResult(new List<FrontInspectionResult>());
            bool hasCaptures = captureTickets.Count > 0;
            string returnSummary = null;
            if (Math.Abs(returnDelta) > 0.1)
            {
                try
                {
                    using (CancellationTokenSource moveBackCts = CancellationTokenSource.CreateLinkedTokenSource(token))
                    {
                        moveBackCts.CancelAfter(TimeSpan.FromSeconds(45));
                        Task<string> moveTask = controller.MoveRelativeAsync(returnDelta, moveBackCts.Token);
                        Task<List<FrontInspectionResult>> processingTask = hasCaptures
                            ? ProcessFrontCapturesAsync(captureTickets, token)
                            : Task.FromResult(new List<FrontInspectionResult>());
                        await Task.WhenAll(moveTask, processingTask).ConfigureAwait(false);
                        processTask = processingTask;
                        returnSummary = moveTask.Result;
                        if (!IsDisposed)
                        {
                            BeginInvoke(new MethodInvoker(() =>
                            {
                                outToLog($"[Turntable] Return to home: {returnSummary}", LogStatus.Info);
                            }));
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Turntable] Failed to return to home position: {ex.Message}", LogStatus.Error);
                        }));
                    }
                    return;
                }
            }
            else
            {
                if (hasCaptures)
                {
                    processTask = ProcessFrontCapturesAsync(captureTickets, token);
                    await processTask.ConfigureAwait(false);
                }
                if (!IsDisposed)
                {
                    BeginInvoke(new MethodInvoker(() =>
                    {
                        outToLog("[Turntable] Already at home position.", LogStatus.Info);
                    }));
                }
            }

            List<FrontInspectionResult> inspections = processTask.Result ?? new List<FrontInspectionResult>();
            if (!IsDisposed)
            {
                BeginInvoke(new MethodInvoker(() =>
                {
                    outToLog("[Sequence] Capture sequence completed.", LogStatus.Success);
                    ApplyFrontInspectionResults(sourceImage, inspections);
                }));
            }
        }

        private string EnsureFrontCaptureDirectory(DetectImg sourceImage)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (sourceImage != null && !string.IsNullOrWhiteSpace(sourceImage.OriImgPath))
            {
                string candidate = Path.GetDirectoryName(sourceImage.OriImgPath);
                if (!string.IsNullOrEmpty(candidate))
                {
                    baseDirectory = candidate;
                }
            }

            string folderName = sourceImage != null && !string.IsNullOrWhiteSpace(sourceImage.OriImgPath)
                ? Path.GetFileNameWithoutExtension(sourceImage.OriImgPath) + "_Front"
                : "FrontCaptures";

            string directory = Path.Combine(baseDirectory, folderName);
            Directory.CreateDirectory(directory);
            return directory;
        }

        private static string BuildAngleToken(double angle)
        {
            string formatted = angle.ToString("+000.0;-000.0;+000.0", CultureInfo.InvariantCulture);
            formatted = formatted.Replace("+", "P").Replace("-", "N").Replace(".", "_");
            return $"Angle_{formatted}";
        }

        private async Task<List<FrontInspectionResult>> ProcessFrontCapturesAsync(List<FrontCaptureTicket> captures, CancellationToken token)
        {
            if (captures == null || captures.Count == 0 || defectContext?.Process == null)
            {
                return new List<FrontInspectionResult>();
            }

            List<FrontInspectionResult> results = new List<FrontInspectionResult>(captures.Count);
            await Task.Run(() =>
            {
                foreach (FrontCaptureTicket ticket in captures)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Defect] Processing index {ticket.Sequence}...", LogStatus.Progress);
                        }));

                        List<ObjectInfo> detections;
                        Bitmap overlayBitmap;
                        Bitmap rawBitmap;
                        using (Mat raw = CvInvoke.Imread(ticket.ImagePath, ImreadModes.ColorBgr))
                        {
                            if (raw == null || raw.IsEmpty)
                            {
                                throw new InvalidOperationException("Unable to load captured frame for defect inspection.");
                            }

                            using (Mat clone = raw.Clone())
                            {
                                defectContext.Process.Detect(clone, out detections);
                            }

                            rawBitmap = raw.ToBitmap();
                            using (Mat overlayMat = raw.Clone())
                            {
                                DrawDefectAnnotations(overlayMat, detections);
                                overlayBitmap = overlayMat.ToBitmap();
                            }
                        }

                        FrontInspectionResult result = new FrontInspectionResult(ticket.Sequence, ticket.ImagePath, ticket.AngleDegrees, ticket.CapturedAt, detections, overlayBitmap, rawBitmap);
                        results.Add(result);
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            string summary = result.BuildSummary();
                            LogStatus status = result.HasDefects ? LogStatus.Warning : LogStatus.Success;
                            outToLog($"[Defect] Index {result.Sequence}: {summary}", status);
                        }));
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Defect] Index {ticket.Sequence} failed: {ex.Message}", LogStatus.Error);
                        }));
                    }
                }
            }, token).ConfigureAwait(false);

            results.Sort((a, b) => a.Sequence.CompareTo(b.Sequence));
            return results;
        }

        private static void DrawDefectAnnotations(Mat image, List<ObjectInfo> detections)
        {
            if (image == null || image.IsEmpty || detections == null || detections.Count == 0)
            {
                return;
            }

            foreach (ObjectInfo obj in detections)
            {
                Rectangle rect = obj.DisplayRec;
                if (rect.Width <= 0 || rect.Height <= 0)
                {
                    continue;
                }

                MCvScalar color = GetDefectColor(obj?.name);
                CvInvoke.Rectangle(image, rect, color, 3);
                string label = string.IsNullOrWhiteSpace(obj?.name) ? "(unknown)" : obj.name;
                CvInvoke.PutText(image, label, new Point(rect.Left, Math.Max(0, rect.Top - 8)), FontFace.HersheySimplex, 0.7, color, 2);
            }
        }

        private static MCvScalar GetDefectColor(string name)
        {
            Color color;
            if (string.IsNullOrWhiteSpace(name))
            {
                color = Color.FromArgb(200, 200, 200);
            }
            else if (!string.IsNullOrEmpty(name) && name.Equals("Fill", StringComparison.OrdinalIgnoreCase))
            {
                color = Color.FromArgb(255, 99, 132);
            }
            else if (!string.IsNullOrEmpty(name) && name.Equals("Bubble", StringComparison.OrdinalIgnoreCase))
            {
                color = Color.FromArgb(54, 162, 235);
            }
            else if (!string.IsNullOrEmpty(name) && name.Equals("Flash", StringComparison.OrdinalIgnoreCase))
            {
                color = Color.FromArgb(255, 206, 86);
            }
            else if (!string.IsNullOrEmpty(name) && name.Equals("Contamination", StringComparison.OrdinalIgnoreCase))
            {
                color = Color.FromArgb(75, 192, 192);
            }
            else
            {
                color = Color.FromArgb(156, 39, 176);
            }

            return new MCvScalar(color.B, color.G, color.R);
        }

        private void UpdateDefectClassNamesFromInspections(IEnumerable<FrontInspectionResult> inspections)
        {
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (inspections != null)
            {
                foreach (FrontInspectionResult result in inspections)
                {
                    if (result?.Detections == null)
                    {
                        continue;
                    }

                    foreach (ObjectInfo detection in result.Detections)
                    {
                        string name = detection?.name;
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name.Trim());
                        }
                    }
                }
            }

            classNameList = names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();

            if (CB_Field != null && CB_Field.SelectedValue is LogicField field && field == LogicField.ClassName)
            {
                UpdateValueSuggestions(LogicField.ClassName, CB_Value?.Text);
            }
        }

        private void ApplyFrontInspectionResults(DetectImg sourceImage, List<FrontInspectionResult> inspections)
        {
            if (sourceImage == null || inspections == null)
            {
                return;
            }

            if (!ReferenceEquals(sourceImage, currentImage))
            {
                foreach (FrontInspectionResult inspection in inspections)
                {
                    inspection?.Dispose();
                }
                sourceImage.FrontInspectionComplete = true;
                return;
            }

            if (sourceImage.FrontInspections != null)
            {
                foreach (FrontInspectionResult oldResult in sourceImage.FrontInspections)
                {
                    oldResult?.Dispose();
                }
                sourceImage.FrontInspections.Clear();
            }

            sourceImage.FrontInspections.AddRange(inspections);
            UpdateDefectClassNamesFromInspections(sourceImage.FrontInspections);

            if (sourceImage.AttachmentPoints != null)
            {
                foreach (AttachmentPointInfo point in sourceImage.AttachmentPoints)
                {
                    point.FrontInspection = inspections.FirstOrDefault(r => r.Sequence == point.Sequence);
                }
            }

            sourceImage.FrontInspectionComplete = true;
            DetermineInitialFrontSelection(sourceImage);
            RefreshTopOverlay();
            RefreshFrontGallery();
            UpdateFrontPreviewAndLedger();
            UpdateLogicEvaluation();
        }

        private void DetermineInitialFrontSelection(DetectImg image)
        {
            if (image?.FrontInspections == null || image.FrontInspections.Count == 0)
            {
                selectedFrontSequence = -1;
                return;
            }

            FrontInspectionResult failing = image.FrontInspections.FirstOrDefault(r => r.HasDefects);
            FrontInspectionResult first = image.FrontInspections.First();
            selectedFrontSequence = (failing ?? first).Sequence;
        }

        private void RefreshTopOverlay()
        {
            RenderAllDetections();
        }

        private void RefreshFrontGallery()
        {
            if (flowFrontGallery == null)
            {
                return;
            }

            flowFrontGallery.SuspendLayout();
            try
            {
                flowFrontGallery.Controls.Clear();
                DetectImg image = currentImage;
                if (image?.FrontInspections == null || image.FrontInspections.Count == 0)
                {
                    return;
                }

                foreach (FrontInspectionResult result in image.FrontInspections)
                {
                    Control card = CreateGalleryCard(result);
                    flowFrontGallery.Controls.Add(card);
                }
            }
            finally
            {
                flowFrontGallery.ResumeLayout();
            }
        }

        private Control CreateGalleryCard(FrontInspectionResult result)
        {
            Button card = new Button
            {
                Tag = result.Sequence,
                Width = 190,
                Height = 90,
                Margin = new Padding(6),
                TextAlign = ContentAlignment.MiddleLeft,
                ImageAlign = ContentAlignment.MiddleRight,
                FlatStyle = FlatStyle.Flat,
                BackColor = result.HasDefects ? Color.FromArgb(255, 235, 238) : Color.FromArgb(232, 245, 233),
                ForeColor = result.HasDefects ? Color.FromArgb(178, 34, 34) : Color.FromArgb(27, 94, 32),
                Text = $"Index {result.Sequence:00}\nAngle {result.AngleDegrees:+0.0;-0.0;+0.0}°\n{result.BuildSummary()}"
            };
            card.FlatAppearance.BorderSize = result.Sequence == selectedFrontSequence ? 2 : 1;
            card.FlatAppearance.BorderColor = result.Sequence == selectedFrontSequence ? Color.FromArgb(33, 150, 243) : Color.FromArgb(189, 189, 189);
            card.Click += FrontGalleryCard_Click;
            return card;
        }

        private void FrontGalleryCard_Click(object sender, EventArgs e)
        {
            if (sender is Control control && control.Tag is int sequence)
            {
                selectedFrontSequence = sequence;
                RefreshTopOverlay();
                RefreshFrontGallery();
                UpdateFrontPreviewAndLedger();
            }
        }

        private void UpdateFrontPreviewAndLedger()
        {
            FrontInspectionResult selected = null;
            DetectImg image = currentImage;
            if (image?.FrontInspections != null && selectedFrontSequence > 0)
            {
                selected = image.FrontInspections.FirstOrDefault(r => r.Sequence == selectedFrontSequence);
            }

            if (LBL_FrontSummary != null)
            {
                if (selected != null)
                {
                    string status = selected.HasDefects ? "FAIL" : "PASS";
                    LBL_FrontSummary.Text = $"Index {selected.Sequence:00} - {status}: {selected.BuildSummary()}";
                    LBL_FrontSummary.ForeColor = selected.HasDefects ? Color.FromArgb(183, 28, 28) : Color.FromArgb(27, 94, 32);
                }
                else
                {
                    LBL_FrontSummary.Text = "No inspection selected.";
                    LBL_FrontSummary.ForeColor = Color.FromArgb(94, 102, 112);
                }
            }

            if (PB_FrontPreview != null)
            {
                Image imageToShow = null;
                if (selected != null)
                {
                    imageToShow = showFrontOverlay ? selected.OverlayImage != null ? (Image)selected.OverlayImage.Clone() : null
                                                   : selected.RawImage != null ? (Image)selected.RawImage.Clone() : null;
                }
                ReplacePictureBoxImage(PB_FrontPreview, imageToShow);
            }

            if (DG_FrontDefects != null)
            {
                DG_FrontDefects.Rows.Clear();
                if (selected != null && selected.Detections != null)
                {
                    int rowIndex = 1;
                    foreach (ObjectInfo detection in selected.Detections)
                    {
                        Rectangle rect = detection.DisplayRec;
                        float confidence = detection.confidence != 0 ? detection.confidence : detection.classifyScore;
                        double area = Math.Abs(rect.Width) * Math.Abs(rect.Height);
                        string bounds = $"{rect.X},{rect.Y},{rect.Width},{rect.Height}";
                        DG_FrontDefects.Rows.Add(
                            rowIndex++,
                            detection.name,
                            confidence.ToString("0.00"),
                            area.ToString("0"),
                            bounds);
                    }
                }
            }

            if (BT_FrontPrev != null && BT_FrontNext != null)
            {
                bool hasInspections = image?.FrontInspections != null && image.FrontInspections.Count > 0;
                if (!hasInspections)
                {
                    BT_FrontPrev.Enabled = BT_FrontNext.Enabled = false;
                }
                else
                {
                    int index = image.FrontInspections.FindIndex(r => r.Sequence == selectedFrontSequence);
                    BT_FrontPrev.Enabled = index > 0;
                    BT_FrontNext.Enabled = index >= 0 && index < image.FrontInspections.Count - 1;
                }
            }
        }

        private void BT_FrontPrev_Click(object sender, EventArgs e)
        {
            NavigateFrontInspection(-1);
        }

        private void BT_FrontNext_Click(object sender, EventArgs e)
        {
            NavigateFrontInspection(1);
        }

        private void NavigateFrontInspection(int delta)
        {
            DetectImg image = currentImage;
            if (image?.FrontInspections == null || image.FrontInspections.Count == 0)
            {
                return;
            }

            int index = image.FrontInspections.FindIndex(r => r.Sequence == selectedFrontSequence);
            if (index < 0)
            {
                index = 0;
            }

            int newIndex = Math.Max(0, Math.Min(image.FrontInspections.Count - 1, index + delta));
            selectedFrontSequence = image.FrontInspections[newIndex].Sequence;
            RefreshTopOverlay();
            RefreshFrontGallery();
            UpdateFrontPreviewAndLedger();
        }

        private void CHK_ShowOverlay_CheckedChanged(object sender, EventArgs e)
        {
            showFrontOverlay = CHK_ShowOverlay.Checked;
            UpdateFrontPreviewAndLedger();
        }

        private void RecurDrawRec(Mat img, ObjectInfo obj, MCvScalar? color = null)
        {
            if (obj == null)
            {
                return;
            }

            MCvScalar drawColor = color ?? new MCvScalar(0, 255, 0);

            if (obj.name != "twoPointCrop")
            {
                if (obj.affineAngle == 0)
                {
                    CvInvoke.Rectangle(img, obj.DisplayRec, drawColor, 3);
                }
                else
                {
                    PointF p0 = new PointF(obj.DisplayRec.Left, obj.DisplayRec.Top);
                    PointF p1 = TransferPoint(p0, new PointF(obj.DisplayRec.Right, obj.DisplayRec.Top), obj.affineAngle);
                    PointF p2 = TransferPoint(p0, new PointF(obj.DisplayRec.Right, obj.DisplayRec.Bottom), obj.affineAngle);
                    PointF p3 = TransferPoint(p0, new PointF(obj.DisplayRec.Left, obj.DisplayRec.Bottom), obj.affineAngle);
                    CvInvoke.Line(img, Point.Round(p0), Point.Round(p1), drawColor, 3);
                    CvInvoke.Line(img, Point.Round(p1), Point.Round(p2), drawColor, 3);
                    CvInvoke.Line(img, Point.Round(p2), Point.Round(p3), drawColor, 3);
                    CvInvoke.Line(img, Point.Round(p3), Point.Round(p0), drawColor, 3);
                }

                CvInvoke.PutText(img, obj.name, new Point(obj.DisplayRec.X, obj.DisplayRec.Y - 15), FontFace.HersheySimplex, 1.0, new MCvScalar(0, 0, 255), 2);
            }

            if (obj.childObjs != null && obj.childObjs.Count > 0)
            {
                foreach (ObjectInfo child in obj.childObjs)
                {
                    RecurDrawRec(img, child, color);
                }
            }
        }

        private PointF RotatePoint(PointF subPoint, double angle)
        {
            double theta = -angle / 180 * Math.PI;
            PointF point = new PointF
            {
                X = subPoint.X * (float)Math.Cos(theta) + subPoint.Y * (float)Math.Sin(theta),
                Y = -subPoint.X * (float)Math.Sin(theta) + subPoint.Y * (float)Math.Cos(theta)
            };
            return point;
        }

        private PointF TransferPoint(PointF oriPoint, PointF subPoint, double angle)
        {
            PointF distanceVector = RotatePoint(new PointF(subPoint.X - oriPoint.X, subPoint.Y - oriPoint.Y), angle);
            return new PointF(oriPoint.X + distanceVector.X, oriPoint.Y + distanceVector.Y);
        }

        private void outToLog(string output, LogStatus status = LogStatus.Info)
        {
            string formatted = FormatLogMessage(output, status);
            if (string.IsNullOrEmpty(formatted))
            {
                return;
            }

            if (RTB_Info.TextLength > 0)
            {
                RTB_Info.AppendText(Environment.NewLine);
            }
            RTB_Info.AppendText(formatted);
            RTB_Info.ScrollToCaret();
        }

        private string FormatLogMessage(string rawMessage, LogStatus status)
        {
            if (string.IsNullOrWhiteSpace(rawMessage))
            {
                return string.Empty;
            }

            string text = rawMessage.Trim();
            string description = null;

            if (text.StartsWith("ConfigParameter=", StringComparison.OrdinalIgnoreCase))
            {
                string config = ExtractAfter(text, "ConfigParameter=");
                string tool = ExtractAfter(text, "DeepLearningTool=");
                description = $"Deep learning configuration loaded (Parameter: {config}, Tool: {tool}).";
            }
            else if (text.StartsWith("Add tool", StringComparison.OrdinalIgnoreCase))
            {
                string index = ExtractAfter(text, "Add tool");
                description = $"Loaded tool index {index}.";
            }
            else if (text.StartsWith("TaskProcess:", StringComparison.OrdinalIgnoreCase))
            {
                string detail = ExtractAfter(text, "TaskProcess:");
                description = $"Task process event: {detail}.";
            }
            else if (text.StartsWith("GPUIndex", StringComparison.OrdinalIgnoreCase))
            {
                string gpu = ExtractAfter(text, "GPUIndex");
                description = $"GPU index selected: {gpu}.";
            }
            else if (text.StartsWith("create_server", StringComparison.OrdinalIgnoreCase))
            {
                string statusValue = ExtractJsonValue(text, "status");
                string port = ExtractJsonValue(text, "port");
                description = $"Inference server started (port {port}, status: {statusValue}).";
            }
            else if (text.StartsWith("CheckPort", StringComparison.OrdinalIgnoreCase))
            {
                Match portMatch = Regex.Match(text, @"CheckPort\s+(?<port>\d+)", RegexOptions.IgnoreCase);
                string port = portMatch.Success ? portMatch.Groups["port"].Value : string.Empty;
                description = string.IsNullOrEmpty(port) ? "Checking inference server port." : $"Verified inference server port {port}.";
            }
            else if (text.StartsWith("[Info] Created new shared memory", StringComparison.OrdinalIgnoreCase))
            {
                string id = ExtractAfter(text, ":");
                description = $"Shared memory created ({id}).";
            }
            else if (text.StartsWith("[Info] Link shared memory result", StringComparison.OrdinalIgnoreCase))
            {
                string statusValue = ExtractJsonValue(text, "status");
                string message = ExtractJsonValue(text, "message");
                description = $"Shared memory link {statusValue}. {message}".Trim();
            }
            else if (text.StartsWith("Init infer model", StringComparison.OrdinalIgnoreCase))
            {
                string statusValue = ExtractJsonValue(text, "status");
                string message = ExtractJsonValue(text, "message");
                string modelName = Regex.Match(message ?? string.Empty, @"Model\s+'(?<name>[^']+)'", RegexOptions.IgnoreCase).Groups["name"].Value;
                if (string.IsNullOrEmpty(modelName))
                {
                    modelName = "unknown";
                }
                description = $"Model initialized ({modelName}, status: {statusValue}).";
            }
            else
            {
                Match taskMatch = Regex.Match(text, @"^-{6}task\s+(?<index>\d+)\s*:\s*(?<duration>\d+)\s*ms", RegexOptions.IgnoreCase);
                if (taskMatch.Success)
                {
                    string idx = taskMatch.Groups["index"].Value;
                    string duration = taskMatch.Groups["duration"].Value;
                    description = $"Task {idx} completed in {duration} ms.";
                }
            }

            if (description == null)
            {
                if (text.StartsWith("SolVision Version", StringComparison.OrdinalIgnoreCase))
                {
                    description = text.Replace("SolVision Version", "SDK version");
                }
                else if (text.IndexOf("Camera is disconnected", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    description = "Camera connection not detected.";
                }
                else if (text.StartsWith("???", StringComparison.OrdinalIgnoreCase))
                {
                    description = $"SDK message (untranslated): {text.Trim('?').Trim()}";
                }
                else
                {
                    description = text;
                }
            }

            if (status == LogStatus.Info && !string.IsNullOrEmpty(description))
            {
                string normalized = description.ToLowerInvariant();
                if (normalized.Contains("failed") || normalized.Contains("error"))
                {
                    status = LogStatus.Error;
                }
                else if (normalized.Contains("not detected") || normalized.Contains("disconnected") || normalized.Contains("missing"))
                {
                    status = LogStatus.Warning;
                }
                else if (normalized.Contains("success") || normalized.Contains("initialized") || normalized.Contains("completed"))
                {
                    status = LogStatus.Success;
                }
            }

            string statusToken;
            switch (status)
            {
                case LogStatus.Success:
                    statusToken = "DONE";
                    break;
                case LogStatus.Warning:
                    statusToken = "WARN";
                    break;
                case LogStatus.Error:
                    statusToken = "FAIL";
                    break;
                case LogStatus.Progress:
                    statusToken = "WORK";
                    break;
                default:
                    statusToken = "INFO";
                    break;
            }

            return $"[{DateTime.Now:HH:mm:ss}][{statusToken}] {description}";
        }

        private static string ExtractAfter(string source, string marker)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(marker))
            {
                return string.Empty;
            }

            int index = source.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                return string.Empty;
            }

            string segment = source.Substring(index + marker.Length).Trim();
            if (segment.StartsWith("="))
            {
                segment = segment.Substring(1).Trim();
            }
            segment = segment.TrimEnd(',');
            return segment;
        }

        private static string ExtractJsonValue(string source, string key)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(key))
            {
                return string.Empty;
            }

            Match match = Regex.Match(source, $"\"{Regex.Escape(key)}\"\\s*:\\s*\"?(?<value>[^\"}}]+)\"?", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string value = match.Groups["value"].Value.Trim();
                value = value.Trim('\'');
                return value;
            }

            return string.Empty;
        }

        private void SetPictureBoxImage(PictureBox pictureBox, Image image)
        {
            if (pictureBox == null)
            {
                return;
            }

            Image oldImage = pictureBox.Image;
            pictureBox.Image = image;
            oldImage?.Dispose();
        }

        private void ApplyTheme()
        {
            Font baseFont = new Font("Segoe UI", 9F);
            this.Font = baseFont;

            Color background = Color.FromArgb(246, 248, 252);
            Color surface = Color.White;
            Color border = Color.FromArgb(223, 227, 235);
            Color accentPrimary = Color.FromArgb(37, 99, 235);
            Color accentHover = Color.FromArgb(53, 110, 240);
            Color accentSecondary = Color.FromArgb(94, 102, 112);

            this.BackColor = background;

            tableLayoutPanelMain.BackColor = background;
            tableLayoutSteps.BackColor = background;
            tableLayoutImages.BackColor = background;

            StyleGroupSurface(groupStep1, surface, border);
            StyleGroupSurface(groupStep2, surface, border);
            StyleGroupSurface(groupStep4, surface, border);
            StyleGroupSurface(groupStep5, surface, border);
            StyleGroupSurface(groupGallery, surface, border);
            StyleGroupSurface(groupStep7, surface, border);
            StyleGroupSurface(groupStep3, surface, border);
            StyleGroupSurface(groupStep6, surface, border);
            StyleGroupSurface(groupDefectLedger, surface, border);
            StyleGroupSurface(groupLogic, surface, border);
            if (groupInitAttachment != null)
            {
                StyleGroupSurface(groupInitAttachment, surface, border);
            }
            if (groupInitDefect != null)
            {
                StyleGroupSurface(groupInitDefect, surface, border);
            }
            if (groupInitTurntable != null)
            {
                StyleGroupSurface(groupInitTurntable, surface, border);
            }

            if (leftTabs != null)
            {
                leftTabs.Font = new Font("Segoe UI Semibold", 9F);
                leftTabs.Appearance = TabAppearance.Normal;
                leftTabs.DrawMode = TabDrawMode.Normal;
                leftTabs.ItemSize = new Size(120, 28);
                leftTabs.SizeMode = TabSizeMode.Fixed;
                foreach (TabPage tab in leftTabs.TabPages)
                {
                    tab.BackColor = surface;
                    tab.ForeColor = Color.FromArgb(58, 66, 78);
                }
            }

            StylePrimaryButton(BT_LoadProject, accentPrimary, accentHover);
            StylePrimaryButton(BT_LoadImg, accentPrimary, accentHover);
            StylePrimaryButton(BT_Detect, Color.FromArgb(18, 163, 123), Color.FromArgb(16, 149, 113));
            StylePrimaryButton(BT_FrontPrev, accentSecondary, Color.FromArgb(79, 88, 99));
            StylePrimaryButton(BT_FrontNext, accentSecondary, Color.FromArgb(79, 88, 99));
            BT_FrontPrev.Height = 28;
            BT_FrontNext.Height = 28;

            foreach (Button initButton in new[] { BT_InitAttachmentBrowse, BT_InitAttachmentLoad, BT_InitDefectBrowse, BT_InitDefectLoad, BT_TurntableRefresh, BT_TurntableConnect, BT_TurntableHome })
            {
                if (initButton != null)
                {
                    StylePrimaryButton(initButton, accentPrimary, accentHover);
                    initButton.Height = 32;
                    initButton.Margin = new Padding(4);
                }
            }

            foreach (TextBox initPath in new[] { TB_InitAttachmentPath, TB_InitDefectPath })
            {
                if (initPath != null)
                {
                    initPath.BorderStyle = BorderStyle.FixedSingle;
                    initPath.BackColor = Color.FromArgb(252, 253, 255);
                    initPath.ForeColor = accentSecondary;
                }
            }

            TB_ProjectPath.BorderStyle = BorderStyle.FixedSingle;
            TB_ProjectPath.BackColor = Color.FromArgb(252, 253, 255);
            TB_ProjectPath.ForeColor = accentSecondary;

            RTB_Info.BorderStyle = BorderStyle.None;
            RTB_Info.BackColor = Color.FromArgb(249, 250, 253);
            RTB_Info.ForeColor = accentSecondary;
            RTB_Info.Font = new Font("Consolas", 9F);
            RTB_Info.Margin = new Padding(16);
            flowFrontGallery.BackColor = surface;

            ConfigureDataGridView(DG_Detections, surface, border, accentPrimary, accentSecondary);
            ConfigureDataGridView(DG_FrontDefects, surface, border, accentPrimary, accentSecondary);

            PB_OriginalImage.BackColor = Color.FromArgb(245, 247, 252);
            PB_OriginalImage.BorderStyle = BorderStyle.FixedSingle;
            PB_FrontPreview.BackColor = Color.FromArgb(245, 247, 252);
            PB_FrontPreview.BorderStyle = BorderStyle.FixedSingle;

            tableLayoutLogic.BackColor = surface;
            panelLogicEditor.BackColor = surface;
            tableLayoutLogicButtons.BackColor = Color.Transparent;

            TV_Logic.BackColor = Color.FromArgb(249, 250, 253);
            TV_Logic.ForeColor = accentSecondary;
            TV_Logic.BorderStyle = BorderStyle.None;
            TV_Logic.Font = new Font("Segoe UI", 9F);
            TV_Logic.LineColor = Color.FromArgb(210, 216, 226);

            foreach (ComboBox combo in new[] { CB_GroupOperator, CB_Field, CB_Operator, CB_Value, CB_TurntablePort })
            {
                combo.FlatStyle = FlatStyle.Flat;
                combo.BackColor = Color.FromArgb(252, 253, 255);
                combo.ForeColor = accentSecondary;
                combo.Margin = new Padding(0, 4, 0, 4);
            }

            CB_Value.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            CB_Value.AutoCompleteSource = AutoCompleteSource.ListItems;
            CB_Value.DropDownStyle = ComboBoxStyle.DropDown;

            StylePrimaryButton(BT_AddRule, accentPrimary, accentHover);
            StylePrimaryButton(BT_AddGroup, accentPrimary, accentHover);
            BT_AddRule.Height = 28;
            BT_AddGroup.Height = 28;
            BT_AddRule.Margin = new Padding(2, 0, 2, 0);
            BT_AddGroup.Margin = new Padding(2, 0, 2, 0);

            StylePrimaryButton(BT_LogicHelp, accentSecondary, Color.FromArgb(76, 84, 94));
            BT_LogicHelp.BackColor = Color.FromArgb(231, 233, 238);
            BT_LogicHelp.ForeColor = Color.FromArgb(58, 66, 78);
            BT_LogicHelp.Width = 28;
            BT_LogicHelp.Height = 28;

            StyleDangerButton(BT_RemoveNode, Color.FromArgb(224, 68, 67), Color.FromArgb(205, 58, 57));
            BT_RemoveNode.Height = 28;
            BT_RemoveNode.Margin = new Padding(2, 0, 2, 0);

            StyleStatusLabel(LBL_ResultStatus);
            StyleStatusLabel(LBL_WorkflowStatus);
            StyleStatusLabel(LBL_InitSummary);
            StyleStatusLabel(LBL_TurntableStatus);
            SetLogicStatusNeutral("Awaiting defect inspection.");

            this.Padding = new Padding(12);
        }

        private void BuildLogicBuilderUI()
        {
            toolTip = new ToolTip();

            leftTabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Appearance = TabAppearance.Normal,
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(120, 28)
            };

            tabInitialize = new TabPage("Initialize")
            {
                Padding = new Padding(10)
            };
            tabWorkflow = new TabPage("Workflow")
            {
                Padding = new Padding(8)
            };
            tabLogicBuilder = new TabPage("Logic Builder")
            {
                Padding = new Padding(10)
            };

initLayout = new TableLayoutPanel
{
    Dock = DockStyle.Fill,
    ColumnCount = 1,
    RowCount = 5
};
initLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 25F));
initLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 25F));
initLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 25F));
initLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 25F));
initLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));

groupInitAttachment = new GroupBox
{
    Text = "Attachment Model",
    Dock = DockStyle.Fill,
    Padding = new Padding(12, 20, 12, 12)
};

TableLayoutPanel attachmentLayout = new TableLayoutPanel
{
    ColumnCount = 3,
    Dock = DockStyle.Fill,
    Margin = new Padding(0)
};
attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
attachmentLayout.RowCount = 3;
attachmentLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
attachmentLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
attachmentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

TB_InitAttachmentPath = new TextBox
{
    Dock = DockStyle.Fill,
    ReadOnly = true,
    Margin = new Padding(0, 0, 0, 8)
};
attachmentLayout.Controls.Add(TB_InitAttachmentPath, 0, 0);
attachmentLayout.SetColumnSpan(TB_InitAttachmentPath, 3);

BT_InitAttachmentBrowse = new Button
{
    Text = "Browse...",
    Dock = DockStyle.Fill
};
BT_InitAttachmentBrowse.Click += BT_InitAttachmentBrowse_Click;
attachmentLayout.Controls.Add(BT_InitAttachmentBrowse, 1, 1);

BT_InitAttachmentLoad = new Button
{
    Text = "Load",
    Dock = DockStyle.Fill,
    Enabled = false
};
BT_InitAttachmentLoad.Click += BT_InitAttachmentLoad_Click;
attachmentLayout.Controls.Add(BT_InitAttachmentLoad, 2, 1);

LBL_InitAttachmentStatus = new Label
{
    Text = "Not loaded.",
    Dock = DockStyle.Fill,
    TextAlign = ContentAlignment.MiddleLeft,
    Padding = new Padding(8, 6, 8, 6)
};
attachmentLayout.Controls.Add(LBL_InitAttachmentStatus, 0, 2);
attachmentLayout.SetColumnSpan(LBL_InitAttachmentStatus, 3);

groupInitAttachment.Controls.Add(attachmentLayout);

groupInitDefect = new GroupBox
{
    Text = "Defect Model",
    Dock = DockStyle.Fill,
    Padding = new Padding(12, 20, 12, 12)
};

TableLayoutPanel defectLayout = new TableLayoutPanel
{
    ColumnCount = 3,
    Dock = DockStyle.Fill,
    Margin = new Padding(0)
};
defectLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
defectLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
defectLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
defectLayout.RowCount = 3;
defectLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
defectLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
defectLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

TB_InitDefectPath = new TextBox
{
    Dock = DockStyle.Fill,
    ReadOnly = true,
    Margin = new Padding(0, 0, 0, 8)
};
defectLayout.Controls.Add(TB_InitDefectPath, 0, 0);
defectLayout.SetColumnSpan(TB_InitDefectPath, 3);

BT_InitDefectBrowse = new Button
{
    Text = "Browse...",
    Dock = DockStyle.Fill
};
BT_InitDefectBrowse.Click += BT_InitDefectBrowse_Click;
defectLayout.Controls.Add(BT_InitDefectBrowse, 1, 1);

BT_InitDefectLoad = new Button
{
    Text = "Load",
    Dock = DockStyle.Fill,
    Enabled = false
};
BT_InitDefectLoad.Click += BT_InitDefectLoad_Click;
defectLayout.Controls.Add(BT_InitDefectLoad, 2, 1);

LBL_InitDefectStatus = new Label
{
    Text = "Not loaded.",
    Dock = DockStyle.Fill,
    TextAlign = ContentAlignment.MiddleLeft,
    Padding = new Padding(8, 6, 8, 6)
};
defectLayout.Controls.Add(LBL_InitDefectStatus, 0, 2);
defectLayout.SetColumnSpan(LBL_InitDefectStatus, 3);

groupInitDefect.Controls.Add(defectLayout);

groupInitCameras = new GroupBox
{
    Text = "Cameras",
    Dock = DockStyle.Fill,
    Padding = new Padding(12, 20, 12, 12)
};

TableLayoutPanel camerasLayout = new TableLayoutPanel
{
    ColumnCount = 1,
    Dock = DockStyle.Fill,
    Margin = new Padding(0),
    RowCount = 2
};
camerasLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
camerasLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

FlowLayoutPanel camerasToolbar = new FlowLayoutPanel
{
    Dock = DockStyle.Fill,
    FlowDirection = FlowDirection.LeftToRight,
    WrapContents = false,
    Margin = new Padding(0),
    Padding = new Padding(0)
};

BT_CameraRefresh = new Button
{
    Text = "Refresh",
    AutoSize = true
};
BT_CameraRefresh.Click += BT_CameraRefresh_Click;
camerasToolbar.Controls.Add(BT_CameraRefresh);

Label cameraToolbarHint = new Label
{
    AutoSize = true,
    Margin = new Padding(12, 8, 0, 0),
    Text = "Pick the top and front cameras, then use Capture Preview to refresh the panels on the right."
};
camerasToolbar.Controls.Add(cameraToolbarHint);

camerasLayout.Controls.Add(camerasToolbar, 0, 0);

TableLayoutPanel cameraColumns = new TableLayoutPanel
{
    ColumnCount = 2,
    Dock = DockStyle.Fill,
    Margin = new Padding(0)
};
cameraColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
cameraColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

Label LBL_TopCameraDetail;
TableLayoutPanel topCameraLayout = CreateCameraRoleLayout(
    topCameraContext,
    "Top Camera",
    out CB_TopCameraSelect,
    out BT_TopCameraConnect,
    out BT_TopCameraCapture,
    out LBL_TopCameraDetail,
    out LBL_TopCameraStatus);

Label LBL_FrontCameraDetail;
TableLayoutPanel frontCameraLayout = CreateCameraRoleLayout(
    frontCameraContext,
    "Front Camera",
    out CB_FrontCameraSelect,
    out BT_FrontCameraConnect,
    out BT_FrontCameraCapture,
    out LBL_FrontCameraDetail,
    out LBL_FrontCameraStatus);

cameraColumns.Controls.Add(topCameraLayout, 0, 0);
cameraColumns.Controls.Add(frontCameraLayout, 1, 0);

camerasLayout.Controls.Add(cameraColumns, 0, 1);
groupInitCameras.Controls.Add(camerasLayout);

topCameraContext.Selector = CB_TopCameraSelect;
topCameraContext.ConnectButton = BT_TopCameraConnect;
topCameraContext.CaptureButton = BT_TopCameraCapture;
topCameraContext.StatusLabel = LBL_TopCameraStatus;
topCameraContext.DetailLabel = LBL_TopCameraDetail;

frontCameraContext.Selector = CB_FrontCameraSelect;
frontCameraContext.ConnectButton = BT_FrontCameraConnect;
frontCameraContext.CaptureButton = BT_FrontCameraCapture;
frontCameraContext.StatusLabel = LBL_FrontCameraStatus;
frontCameraContext.DetailLabel = LBL_FrontCameraDetail;

CB_TopCameraSelect.SelectedIndexChanged += (s, e) => UpdateCameraSelectionChanged(topCameraContext);
CB_FrontCameraSelect.SelectedIndexChanged += (s, e) => UpdateCameraSelectionChanged(frontCameraContext);
BT_TopCameraConnect.Click += (s, e) => ToggleCameraConnection(topCameraContext);
BT_FrontCameraConnect.Click += (s, e) => ToggleCameraConnection(frontCameraContext);
BT_TopCameraCapture.Click += (s, e) => CaptureCameraPreview(topCameraContext);
BT_FrontCameraCapture.Click += (s, e) => CaptureCameraPreview(frontCameraContext);

RefreshCameraUiState(topCameraContext);
RefreshCameraUiState(frontCameraContext);
RefreshCameraList();

groupInitTurntable = new GroupBox
{
    Text = "Turntable",
    Dock = DockStyle.Fill,
    Padding = new Padding(12, 20, 12, 12)
};

TableLayoutPanel turntableLayout = new TableLayoutPanel
{
    ColumnCount = 3,
    Dock = DockStyle.Fill,
    Margin = new Padding(0)
};
turntableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
turntableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
turntableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
turntableLayout.RowCount = 3;
turntableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
turntableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
turntableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

CB_TurntablePort = new ComboBox
{
    Dock = DockStyle.Fill,
    DropDownStyle = ComboBoxStyle.DropDownList
};
turntableLayout.Controls.Add(CB_TurntablePort, 0, 0);

BT_TurntableRefresh = new Button
{
    Text = "Refresh",
    Dock = DockStyle.Fill
};
BT_TurntableRefresh.Click += BT_TurntableRefresh_Click;
turntableLayout.Controls.Add(BT_TurntableRefresh, 1, 0);

BT_TurntableConnect = new Button
{
    Text = "Connect",
    Dock = DockStyle.Fill
};
BT_TurntableConnect.Click += BT_TurntableConnect_Click;
turntableLayout.Controls.Add(BT_TurntableConnect, 2, 0);

Panel turntableSpacer = new Panel { Dock = DockStyle.Fill };
turntableLayout.Controls.Add(turntableSpacer, 0, 1);
turntableLayout.SetColumnSpan(turntableSpacer, 2);

BT_TurntableHome = new Button
{
    Text = "Home",
    Dock = DockStyle.Fill,
    Enabled = false
};
BT_TurntableHome.Click += BT_TurntableHome_Click;
turntableLayout.Controls.Add(BT_TurntableHome, 2, 1);

LBL_TurntableStatus = new Label
{
    Text = "Disconnected.",
    Dock = DockStyle.Fill,
    TextAlign = ContentAlignment.MiddleLeft,
    Padding = new Padding(8, 6, 8, 6)
};
turntableLayout.Controls.Add(LBL_TurntableStatus, 0, 2);
turntableLayout.SetColumnSpan(LBL_TurntableStatus, 3);

groupInitTurntable.Controls.Add(turntableLayout);

LBL_InitSummary = new Label
{
    Dock = DockStyle.Fill,
    TextAlign = ContentAlignment.MiddleLeft,
    Padding = new Padding(12, 10, 12, 10),
    Margin = new Padding(0, 12, 0, 0)
};

initLayout.Controls.Add(groupInitAttachment, 0, 0);
initLayout.Controls.Add(groupInitDefect, 0, 1);
initLayout.Controls.Add(groupInitCameras, 0, 2);
initLayout.Controls.Add(groupInitTurntable, 0, 3);
initLayout.Controls.Add(LBL_InitSummary, 0, 4);

tabInitialize.Controls.Add(initLayout);


            tableLayoutPanelMain.Controls.Remove(tableLayoutSteps);
            tableLayoutSteps.Dock = DockStyle.Fill;
            workflowLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            workflowLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            workflowLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
            workflowLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            LBL_WorkflowStatus = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Awaiting project.",
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(16, 12, 16, 12),
                Margin = new Padding(0, 8, 0, 0)
            };

            workflowLayout.Controls.Add(tableLayoutSteps, 0, 0);
            workflowLayout.Controls.Add(LBL_WorkflowStatus, 0, 1);
            tabWorkflow.Controls.Add(workflowLayout);

            groupLogic = new GroupBox
            {
                Text = "Logic Rules",
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 24, 12, 12)
            };

            tableLayoutLogic = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            tableLayoutLogic.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            tableLayoutLogic.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
            tableLayoutLogic.RowCount = 2;
            tableLayoutLogic.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutLogic.RowStyles.Add(new RowStyle(SizeType.Absolute, 56F));

            TV_Logic = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            TV_Logic.AfterSelect += TV_Logic_AfterSelect;

            panelLogicEditor = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };

            tableLayoutLogicEditor = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            tableLayoutLogicEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            tableLayoutLogicEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62F));
            tableLayoutLogicEditor.RowCount = 6;
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            labelLogicGroup = new Label
            {
                Text = "Group",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            CB_GroupOperator = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            CB_GroupOperator.SelectedIndexChanged += CB_GroupOperator_SelectedIndexChanged;

            labelLogicField = new Label
            {
                Text = "Field",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            CB_Field = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            CB_Field.SelectedIndexChanged += CB_Field_SelectedIndexChanged;

            labelLogicOperator = new Label
            {
                Text = "Operator",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            CB_Operator = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            CB_Operator.SelectedIndexChanged += CB_Operator_SelectedIndexChanged;

            labelLogicValue = new Label
            {
                Text = "Value",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            CB_Value = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            CB_Value.SelectedIndexChanged += CB_Value_TextChanged;
            CB_Value.TextChanged += CB_Value_TextChanged;
            CB_Value.Margin = new Padding(0, 0, 0, 10);

            tableLayoutLogicButtons = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                Margin = new Padding(0, 10, 0, 4)
            };
            tableLayoutLogicButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34F));
            tableLayoutLogicButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33F));
            tableLayoutLogicButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33F));

            BT_AddRule = new Button
            {
                Text = "Add Fail Rule",
                Dock = DockStyle.Fill
            };
            BT_AddRule.Click += BT_AddRule_Click;

            BT_AddGroup = new Button
            {
                Text = "Add Fail Group",
                Dock = DockStyle.Fill
            };
            BT_AddGroup.Click += BT_AddGroup_Click;

            BT_RemoveNode = new Button
            {
                Text = "Remove",
                Dock = DockStyle.Fill
            };
            BT_RemoveNode.Click += BT_RemoveNode_Click;

            tableLayoutLogicButtons.Controls.Add(BT_AddRule, 0, 0);
            tableLayoutLogicButtons.Controls.Add(BT_AddGroup, 1, 0);
            tableLayoutLogicButtons.Controls.Add(BT_RemoveNode, 2, 0);

            tableLayoutLogicEditor.Controls.Add(labelLogicGroup, 0, 0);
            tableLayoutLogicEditor.Controls.Add(CB_GroupOperator, 1, 0);
            tableLayoutLogicEditor.Controls.Add(labelLogicField, 0, 1);
            tableLayoutLogicEditor.Controls.Add(CB_Field, 1, 1);
            tableLayoutLogicEditor.Controls.Add(labelLogicOperator, 0, 2);
            tableLayoutLogicEditor.Controls.Add(CB_Operator, 1, 2);
            tableLayoutLogicEditor.Controls.Add(labelLogicValue, 0, 3);
            tableLayoutLogicEditor.Controls.Add(CB_Value, 1, 3);
            tableLayoutLogicEditor.Controls.Add(tableLayoutLogicButtons, 0, 4);
            tableLayoutLogicEditor.SetColumnSpan(tableLayoutLogicButtons, 2);
            Panel logicSpacer = new Panel { Dock = DockStyle.Fill };
            tableLayoutLogicEditor.Controls.Add(logicSpacer, 0, 5);
            tableLayoutLogicEditor.SetColumnSpan(logicSpacer, 2);

            panelLogicEditor.Controls.Add(tableLayoutLogicEditor);

            LBL_ResultStatus = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Awaiting defect inspection.",
                TextAlign = ContentAlignment.MiddleLeft,
                BorderStyle = BorderStyle.None,
                Margin = new Padding(3, 0, 3, 0)
            };

            tableLayoutLogic.Controls.Add(TV_Logic, 0, 0);
            tableLayoutLogic.Controls.Add(panelLogicEditor, 1, 0);
            tableLayoutLogic.Controls.Add(LBL_ResultStatus, 0, 1);
            tableLayoutLogic.SetColumnSpan(LBL_ResultStatus, 2);

            groupLogic.Controls.Add(tableLayoutLogic);

            logicHeaderPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 0, 0, 8)
            };

            labelLogicHeader = new Label
            {
                Text = "Define FAIL conditions. During inspection, if any rule or group below evaluates to TRUE, the inspection FAILS.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            BT_LogicHelp = new Button
            {
                Text = "?",
                Width = 32,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(12, 0, 0, 0)
            };
            BT_LogicHelp.FlatAppearance.BorderSize = 0;
            BT_LogicHelp.Click += BT_LogicHelp_Click;
            toolTip.SetToolTip(BT_LogicHelp, "Open help on how logic evaluation affects pass/fail results.");
            toolTip.SetToolTip(CB_GroupOperator, "Choose how child rules combine: ANY (fail if any match) or ALL (fail if all match).");
            toolTip.SetToolTip(CB_Field, "Select the detection property this fail rule should examine.");
            toolTip.SetToolTip(CB_Operator, "Pick how to compare the detection value.");
            toolTip.SetToolTip(CB_Value, "Enter the value that should trigger a fail when the condition is true.");

            logicHeaderPanel.Controls.Add(BT_LogicHelp);
            logicHeaderPanel.Controls.Add(labelLogicHeader);
            BT_LogicHelp.Dock = DockStyle.Right;
            labelLogicHeader.Dock = DockStyle.Fill;

            logicRootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2
            };
            logicRootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56F));
            logicRootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            logicRootLayout.Controls.Add(logicHeaderPanel, 0, 0);
            logicRootLayout.Controls.Add(groupLogic, 0, 1);

            tabLogicBuilder.Controls.Add(logicRootLayout);

            leftTabs.TabPages.Add(tabInitialize);
            leftTabs.TabPages.Add(tabWorkflow);
            leftTabs.TabPages.Add(tabLogicBuilder);

            tableLayoutPanelMain.Controls.Add(leftTabs, 0, 0);
            tableLayoutPanelMain.SetColumn(leftTabs, 0);
            tableLayoutPanelMain.SetRow(leftTabs, 0);
        }

        private void InitializeLoadingIndicator()
        {
            loadingSpinner = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 35,
                Size = new Size(80, 8),
                Visible = false
            };

            loadingIndicatorPanel = new Panel
            {
                AutoSize = true,
                BackColor = Color.Transparent,
                Visible = false
            };

            loadingIndicatorPanel.Controls.Add(loadingSpinner);
            loadingIndicatorPanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            Controls.Add(loadingIndicatorPanel);
            loadingIndicatorPanel.BringToFront();

            Resize += (s, e) => PositionLoadingIndicator();
            Shown += (s, e) => PositionLoadingIndicator();
        }

        private void StyleGroupSurface(Control container, Color surface, Color border)
        {
            if (container == null)
            {
                return;
            }

            if (container is GroupBox groupBox)
            {
                groupBox.BackColor = surface;
                groupBox.ForeColor = Color.FromArgb(60, 72, 88);
                groupBox.Font = new Font("Segoe UI Semibold", 9.5F);
                groupBox.Padding = new Padding(12, 24, 12, 12);
            }
            else
            {
                container.BackColor = surface;
            }

            if (container is TableLayoutPanel panel)
            {
                panel.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
                panel.Padding = new Padding(0);
            }

            container.Margin = new Padding(8);
            container.ForeColor = Color.FromArgb(40, 48, 56);

            if (container is Control ctrl)
            {
                ctrl.Paint += (s, e) =>
                {
                    if (s is GroupBox gb)
                    {
                        Rectangle rect = new Rectangle(0, 7, gb.Width - 1, gb.Height - 8);
                        e.Graphics.Clear(ctrl.Parent.BackColor);
                        using (Brush bodyBrush = new SolidBrush(gb.BackColor))
                        {
                            e.Graphics.FillRectangle(bodyBrush, rect);
                        }
                        using (Pen pen = new Pen(border))
                        {
                            e.Graphics.DrawRectangle(pen, rect);
                        }

                        using (Brush brush = new SolidBrush(gb.BackColor))
                        {
                            SizeF textSize = e.Graphics.MeasureString(gb.Text, gb.Font);
                            RectangleF textRect = new RectangleF(12, 0, textSize.Width + 6, textSize.Height);
                            e.Graphics.FillRectangle(brush, textRect);
                            using (Brush textBrush = new SolidBrush(gb.ForeColor))
                            {
                                e.Graphics.DrawString(gb.Text, gb.Font, textBrush, 12, 0);
                            }
                        }
                    }
                };
            }
        }

        private void StylePrimaryButton(Button button, Color accent, Color hoverAccent)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = accent;
            button.ForeColor = Color.White;
            button.Font = new Font("Segoe UI Semibold", 9.5F);
            button.Height = 38;
            button.Margin = new Padding(6);
            button.Cursor = Cursors.Hand;

            button.MouseEnter += (s, e) => button.BackColor = hoverAccent;
            button.MouseLeave += (s, e) => button.BackColor = accent;
        }

        private void StyleDangerButton(Button button, Color accent, Color hoverAccent)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = accent;
            button.ForeColor = Color.White;
            button.Font = new Font("Segoe UI", 9F);
            button.Cursor = Cursors.Hand;

            button.MouseEnter += (s, e) => button.BackColor = hoverAccent;
            button.MouseLeave += (s, e) => button.BackColor = accent;
        }

        private void StyleStatusLabel(Label label)
        {
            if (label == null)
            {
                return;
            }

            label.BorderStyle = BorderStyle.None;
            label.Padding = new Padding(16, 10, 16, 10);
            label.Margin = new Padding(8, 8, 8, 0);
            label.Font = new Font("Segoe UI Semibold", 9.5F);
            label.BackColor = statusNeutralBackground;
            label.ForeColor = statusNeutralForeground;
        }

        private void UpdateStatusLabel(Label label, Color backColor, Color foreColor, string text)
        {
            if (label == null)
            {
                return;
            }

            label.BackColor = backColor;
            label.ForeColor = foreColor;
            label.Text = text;
        }

        private void PositionLoadingIndicator()
        {
            if (loadingIndicatorPanel == null)
            {
                return;
            }

            int margin = 16;
            int x = ClientSize.Width - loadingIndicatorPanel.Width - margin;
            int y = margin;

            if (x < margin)
            {
                x = margin;
            }

            loadingIndicatorPanel.Location = new Point(x, y);
        }

        private void ConfigureDataGridView(DataGridView grid, Color surface, Color border, Color accent, Color textColor)
        {
            if (grid == null)
            {
                return;
            }

            grid.BackgroundColor = surface;
            grid.GridColor = border;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 243, 248);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(58, 66, 78);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9F);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            grid.ColumnHeadersHeight = 34;

            grid.DefaultCellStyle.BackColor = surface;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(223, 230, 253);
            grid.DefaultCellStyle.SelectionForeColor = accent;
            grid.DefaultCellStyle.ForeColor = textColor;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            grid.RowTemplate.Height = 36;

            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 251, 253);
            grid.RowHeadersVisible = false;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        }

        private void InitializeLogicBuilder()
        {
            logicRoot = new RuleGroupNode { Operator = LogicalOperatorType.Or };

            suppressLogicEvents = true;

            CB_GroupOperator.DisplayMember = "Display";
            CB_GroupOperator.ValueMember = "Value";
            CB_GroupOperator.DataSource = GetGroupOperatorOptions();

            CB_Field.DisplayMember = "Display";
            CB_Field.ValueMember = "Value";
            CB_Field.DataSource = GetFieldOptions();
            CB_Field.SelectedValue = LogicField.ClassName;

            UpdateOperatorOptions(LogicField.ClassName, LogicOperator.Equals);
            UpdateValueSuggestions(LogicField.ClassName);

            suppressLogicEvents = false;

            RefreshLogicTree();
        }

        private List<ComboOption<LogicalOperatorType>> GetGroupOperatorOptions()
        {
            return new List<ComboOption<LogicalOperatorType>>
            {
                new ComboOption<LogicalOperatorType>("Match ANY (OR)", LogicalOperatorType.Or),
                new ComboOption<LogicalOperatorType>("Match ALL (AND)", LogicalOperatorType.And)
            };
        }

        private List<ComboOption<LogicField>> GetFieldOptions()
        {
            return new List<ComboOption<LogicField>>
            {
                new ComboOption<LogicField>("Class Name", LogicField.ClassName),
                new ComboOption<LogicField>("Confidence", LogicField.Confidence),
                new ComboOption<LogicField>("Area (px\u00B2)", LogicField.Area),
                new ComboOption<LogicField>("Detection Count", LogicField.Count)
            };
        }

        private List<ComboOption<LogicOperator>> GetOperatorOptions(LogicField field)
        {
            if (field == LogicField.ClassName)
            {
                return new List<ComboOption<LogicOperator>>
                {
                    new ComboOption<LogicOperator>("equals", LogicOperator.Equals),
                    new ComboOption<LogicOperator>("does not equal", LogicOperator.NotEquals),
                    new ComboOption<LogicOperator>("contains", LogicOperator.Contains),
                    new ComboOption<LogicOperator>("starts with", LogicOperator.StartsWith),
                    new ComboOption<LogicOperator>("ends with", LogicOperator.EndsWith)
                };
            }

            return new List<ComboOption<LogicOperator>>
            {
                new ComboOption<LogicOperator>("=", LogicOperator.Equals),
                new ComboOption<LogicOperator>("?", LogicOperator.NotEquals),
                new ComboOption<LogicOperator>(">", LogicOperator.GreaterThan),
                new ComboOption<LogicOperator>("=", LogicOperator.GreaterThanOrEqual),
                new ComboOption<LogicOperator>("<", LogicOperator.LessThan),
                new ComboOption<LogicOperator>("=", LogicOperator.LessThanOrEqual)
            };
        }

        private void RefreshLogicTree(Guid? selectId = null)
        {
            if (logicRoot == null)
            {
                logicRoot = new RuleGroupNode { Operator = LogicalOperatorType.Or };
            }

            Guid targetId = selectId ?? (TV_Logic.SelectedNode?.Tag as LogicNodeBase)?.Id ?? logicRoot.Id;

            suppressLogicEvents = true;

            TV_Logic.BeginUpdate();
            TV_Logic.Nodes.Clear();
            TreeNode rootNode = CreateTreeNode(logicRoot);
            TV_Logic.Nodes.Add(rootNode);
            rootNode.ExpandAll();
            TreeNode selectedNode = FindNodeById(rootNode, targetId) ?? rootNode;
            TV_Logic.SelectedNode = selectedNode;
            TV_Logic.EndUpdate();

            suppressLogicEvents = false;

            DisplaySelectedNode(TV_Logic.SelectedNode);
            UpdateLogicEvaluation();
        }

        private TreeNode CreateTreeNode(LogicNodeBase node)
        {
            if (node is RuleGroupNode group)
            {
                TreeNode treeNode = new TreeNode(FormatGroupText(group)) { Tag = group };
                foreach (LogicNodeBase child in group.Children)
                {
                    treeNode.Nodes.Add(CreateTreeNode(child));
                }
                return treeNode;
            }

            if (node is RuleConditionNode rule)
            {
                return new TreeNode(FormatRuleText(rule)) { Tag = rule };
            }

            return new TreeNode("(unknown node)") { Tag = node };
        }

        private TreeNode FindNodeById(TreeNode source, Guid id)
        {
            if (source?.Tag is LogicNodeBase logic && logic.Id == id)
            {
                return source;
            }

            foreach (TreeNode child in source.Nodes)
            {
            TreeNode match = FindNodeById(child, id);
            if (match != null)
            {
                return match;
            }
            }

            return null;
        }

        private void DisplaySelectedNode(TreeNode node)
        {
            bool previousSuppress = suppressLogicEvents;
            suppressLogicEvents = true;

            bool isGroup = node?.Tag is RuleGroupNode;
            bool isRule = node?.Tag is RuleConditionNode;

            SetGroupEditorVisibility(isGroup);
            SetRuleEditorVisibility(isRule);

            if (isGroup)
            {
                var group = (RuleGroupNode)node.Tag;
                CB_GroupOperator.SelectedValue = group.Operator;
            }

            if (isRule)
            {
                var rule = (RuleConditionNode)node.Tag;
                CB_Field.SelectedValue = rule.Field;
                UpdateOperatorOptions(rule.Field, rule.Operator);
                UpdateValueSuggestions(rule.Field, rule.Value);
                CB_Value.Text = rule.Value ?? string.Empty;
            }
            else
            {
                CB_Value.Text = string.Empty;
            }

            BT_RemoveNode.Enabled = node != null && node.Tag is LogicNodeBase logicNode && logicNode != logicRoot;

            suppressLogicEvents = previousSuppress;
        }

        private void SetGroupEditorVisibility(bool visible)
        {
            labelLogicGroup.Visible = visible;
            CB_GroupOperator.Visible = visible;
            CB_GroupOperator.Enabled = visible;
        }

        private void SetRuleEditorVisibility(bool visible)
        {
            labelLogicField.Visible = visible;
            CB_Field.Visible = visible;
            labelLogicOperator.Visible = visible;
            CB_Operator.Visible = visible;
            labelLogicValue.Visible = visible;
            CB_Value.Visible = visible;

            CB_Field.Enabled = visible;
            CB_Operator.Enabled = visible;
            CB_Value.Enabled = visible;
        }

        private void TV_Logic_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (suppressLogicEvents)
            {
                return;
            }

            DisplaySelectedNode(e.Node);
        }

        private void CB_GroupOperator_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressLogicEvents)
            {
                return;
            }

            if (TV_Logic.SelectedNode?.Tag is RuleGroupNode group && CB_GroupOperator.SelectedValue is LogicalOperatorType logicOperator)
            {
                group.Operator = logicOperator;
                RefreshLogicTree(group.Id);
            }
        }

        private void CB_Field_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressLogicEvents)
            {
                return;
            }

            if (TV_Logic.SelectedNode?.Tag is RuleConditionNode rule && CB_Field.SelectedValue is LogicField field)
            {
                rule.Field = field;
                rule.Operator = UpdateOperatorOptions(field, rule.Operator);
                UpdateValueSuggestions(field, rule.Value);
                RefreshLogicTree(rule.Id);
            }
        }

        private void CB_Operator_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressLogicEvents)
            {
                return;
            }

            if (TV_Logic.SelectedNode?.Tag is RuleConditionNode rule && CB_Operator.SelectedValue is LogicOperator op)
            {
                rule.Operator = op;
                RefreshLogicTree(rule.Id);
            }
        }

        private void CB_Value_TextChanged(object sender, EventArgs e)
        {
            if (suppressLogicEvents)
            {
                return;
            }

            if (TV_Logic.SelectedNode?.Tag is RuleConditionNode rule)
            {
                rule.Value = CB_Value.Text?.Trim() ?? string.Empty;
                RefreshLogicTree(rule.Id);
            }
        }

        private void BT_AddRule_Click(object sender, EventArgs e)
        {
            var targetGroup = GetTargetGroup();
            var newRule = new RuleConditionNode();
            targetGroup.Children.Add(newRule);
            RefreshLogicTree(newRule.Id);
        }

        private void BT_AddGroup_Click(object sender, EventArgs e)
        {
            var targetGroup = GetTargetGroup();
            var newGroup = new RuleGroupNode();
            targetGroup.Children.Add(newGroup);
            RefreshLogicTree(newGroup.Id);
        }

        private void BT_RemoveNode_Click(object sender, EventArgs e)
        {
            if (TV_Logic.SelectedNode?.Tag is LogicNodeBase node)
            {
                if (node == logicRoot)
                {
                    outToLog("Root group cannot be removed.", LogStatus.Warning);
                    return;
                }

                if (RemoveNodeFromGroup(logicRoot, node.Id))
                {
                    RefreshLogicTree(logicRoot.Id);
                }
            }
        }

        private RuleGroupNode GetTargetGroup()
        {
            if (TV_Logic.SelectedNode?.Tag is RuleGroupNode groupNode)
            {
                return groupNode;
            }

            if (TV_Logic.SelectedNode?.Tag is LogicNodeBase logicNode)
            {
                return FindParentGroup(logicRoot, logicNode.Id) ?? logicRoot;
            }

            return logicRoot ?? (logicRoot = new RuleGroupNode { Operator = LogicalOperatorType.Or });
        }

        private RuleGroupNode FindParentGroup(RuleGroupNode parent, Guid childId)
        {
            foreach (LogicNodeBase child in parent.Children)
            {
                if (child.Id == childId)
                {
                    return parent;
                }

                if (child is RuleGroupNode group)
                {
                    RuleGroupNode candidate = FindParentGroup(group, childId);
                    if (candidate != null)
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private bool RemoveNodeFromGroup(RuleGroupNode parent, Guid targetId)
        {
            for (int i = 0; i < parent.Children.Count; i++)
            {
                LogicNodeBase child = parent.Children[i];
                if (child.Id == targetId)
                {
                    parent.Children.RemoveAt(i);
                    return true;
                }

                if (child is RuleGroupNode group && RemoveNodeFromGroup(group, targetId))
                {
                    if (group.Children.Count == 0)
                    {
                        parent.Children.Remove(group);
                    }
                    return true;
                }
            }

            return false;
        }

        private void UpdateLogicEvaluation()
        {
            if (logicRoot == null)
            {
                return;
            }

            if (logicRoot.Children.Count == 0)
            {
                ApplyStatusToAll("Result: PASS (no rules defined)", statusPassBackground, statusPassForeground);
                return;
            }

            DetectImg image = currentImage;
            if (image == null)
            {
                SetLogicStatusNeutral("Awaiting defect inspections.");
                return;
            }

            if (!image.FrontInspectionComplete)
            {
                SetLogicStatusNeutral("Awaiting defect inspections.");
                return;
            }

            List<ObjectInfo> defectDetections = CollectCurrentDefectDetections();
            bool fail = EvaluateNode(logicRoot, defectDetections, out LogicNodeBase triggered);
            if (fail)
            {
                string description = triggered is RuleConditionNode rule
                    ? FormatRuleText(rule)
                    : triggered is RuleGroupNode group
                        ? FormatGroupText(group)
                        : null;

                string message = description == null ? "Result: FAIL" : $"Result: FAIL ({description})";
                ApplyStatusToAll(message, statusFailBackground, statusFailForeground);
                if (triggered != null)
                {
                    HighlightTriggeredNode(triggered.Id);
                    outToLog($"Fail triggered by: {description}", LogStatus.Warning);
                }
            }
            else
            {
                ApplyStatusToAll("Result: PASS (no defect rules matched)", statusPassBackground, statusPassForeground);
            }
        }

        private void SetLogicStatusNeutral(string message)
        {
            ApplyStatusToAll(message, statusNeutralBackground, statusNeutralForeground);
        }

        private void HighlightTriggeredNode(Guid nodeId)
        {
            if (TV_Logic == null || TV_Logic.Nodes.Count == 0)
            {
                return;
            }

            TreeNode match = FindNodeById(TV_Logic.Nodes[0], nodeId);
            if (match != null)
            {
                suppressLogicEvents = true;
                TV_Logic.SelectedNode = match;
                TV_Logic.Focus();
                suppressLogicEvents = false;
            }
        }

        private void ApplyStatusToAll(string message, Color background, Color foreground)
        {
            UpdateStatusLabel(LBL_ResultStatus, background, foreground, message);
            UpdateStatusLabel(LBL_WorkflowStatus, background, foreground, message);
        }

        private List<ObjectInfo> CollectCurrentDefectDetections()
        {
            DetectImg image = currentImage;
            if (image?.FrontInspections == null || image.FrontInspections.Count == 0)
            {
                return new List<ObjectInfo>();
            }

            List<ObjectInfo> aggregated = new List<ObjectInfo>();
            foreach (FrontInspectionResult inspection in image.FrontInspections)
            {
                if (inspection?.Detections == null || inspection.Detections.Count == 0)
                {
                    continue;
                }

                foreach (ObjectInfo detection in inspection.Detections)
                {
                    if (detection != null)
                    {
                        aggregated.Add(detection);
                    }
                }
            }

            return aggregated;
        }

        private bool EvaluateNode(LogicNodeBase node, List<ObjectInfo> detections, out LogicNodeBase triggered)
        {
            triggered = null;

            if (node is RuleConditionNode rule)
            {
                bool result = EvaluateCondition(rule, detections);
                if (result)
                {
                    triggered = rule;
                }
                return result;
            }

            if (node is RuleGroupNode group)
            {
                if (group.Operator == LogicalOperatorType.And)
                {
                    LogicNodeBase lastTrigger = null;
                    foreach (LogicNodeBase child in group.Children)
                    {
                        if (!EvaluateNode(child, detections, out LogicNodeBase childTrigger))
                        {
                            triggered = null;
                            return false;
                        }

                        if (childTrigger != null)
                        {
                            lastTrigger = childTrigger;
                        }
                    }

                    triggered = lastTrigger ?? group;
                    return group.Children.Count > 0;
                }
                else
                {
                    foreach (LogicNodeBase child in group.Children)
                    {
                        if (EvaluateNode(child, detections, out LogicNodeBase childTrigger))
                        {
                            triggered = childTrigger ?? child;
                            return true;
                        }
                    }

                    triggered = null;
                    return false;
                }
            }

            return false;
        }

        private bool EvaluateCondition(RuleConditionNode rule, List<ObjectInfo> detections)
        {
            string inputValue = rule.Value?.Trim() ?? string.Empty;

            switch (rule.Field)
            {
                case LogicField.ClassName:
                    IEnumerable<string> names = detections.Select(obj => obj.name ?? string.Empty);
                    return EvaluateStringCondition(names, rule.Operator, inputValue);
                case LogicField.Confidence:
                    IEnumerable<double> confidences = detections.Select(GetConfidenceValue);
                    return EvaluateNumericCondition(confidences, rule.Operator, inputValue);
                case LogicField.Area:
                    IEnumerable<double> areas = detections.Select(obj =>
                    {
                        Rectangle rect = obj.DisplayRec;
                        return (double)(Math.Abs(rect.Width) * Math.Abs(rect.Height));
                    });
                    return EvaluateNumericCondition(areas, rule.Operator, inputValue);
                case LogicField.Count:
                    if (!TryParseDouble(inputValue, out double threshold))
                    {
                        return false;
                    }
                    return CompareNumeric(detections.Count, threshold, rule.Operator);
                default:
                    return false;
            }
        }

        private bool EvaluateStringCondition(IEnumerable<string> values, LogicOperator op, string target)
        {
            if (string.IsNullOrWhiteSpace(target) && op != LogicOperator.NotEquals)
            {
                return false;
            }

            List<string> list = values.ToList();

            switch (op)
            {
                case LogicOperator.Equals:
                    return list.Any(v => string.Equals(v, target, StringComparison.OrdinalIgnoreCase));
                case LogicOperator.NotEquals:
                    if (list.Count == 0)
                    {
                        return true;
                    }
                    return list.All(v => !string.Equals(v, target, StringComparison.OrdinalIgnoreCase));
                case LogicOperator.Contains:
                    return list.Any(v => v.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0);
                case LogicOperator.StartsWith:
                    return list.Any(v => v.StartsWith(target, StringComparison.OrdinalIgnoreCase));
                case LogicOperator.EndsWith:
                    return list.Any(v => v.EndsWith(target, StringComparison.OrdinalIgnoreCase));
                default:
                    return false;
            }
        }

        private void BT_LogicHelp_Click(object sender, EventArgs e)
        {
            const string helpMessage =
                "Logic Builder checks detection results after each run.\n\n" +
                "- Groups combine child rules using ALL (AND) or ANY (OR).\n" +
                "- Each rule compares a field (class, confidence, area, count) with the value you provide.\n" +
                "- If ANY rule evaluates to TRUE, the inspection is marked as FAIL.\n" +
                "- With no rules, the inspection defaults to PASS.\n\n" +
                "Tip: Build layered groups to model complex acceptance criteria.";

            MessageBox.Show(helpMessage, "Logic Builder Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool EvaluateNumericCondition(IEnumerable<double> values, LogicOperator op, string target)
        {
            if (!TryParseDouble(target, out double threshold))
            {
                return false;
            }

            List<double> list = values.Where(v => !double.IsNaN(v) && !double.IsInfinity(v)).ToList();
            if (list.Count == 0)
            {
                return false;
            }

            switch (op)
            {
                case LogicOperator.Equals:
                    return list.Any(v => AlmostEqual(v, threshold));
                case LogicOperator.NotEquals:
                    return list.All(v => !AlmostEqual(v, threshold));
                case LogicOperator.GreaterThan:
                    return list.Any(v => v > threshold);
                case LogicOperator.GreaterThanOrEqual:
                    return list.Any(v => v > threshold || AlmostEqual(v, threshold));
                case LogicOperator.LessThan:
                    return list.Any(v => v < threshold);
                case LogicOperator.LessThanOrEqual:
                    return list.Any(v => v < threshold || AlmostEqual(v, threshold));
                default:
                    return false;
            }
        }

        private bool CompareNumeric(double left, double right, LogicOperator op)
        {
            switch (op)
            {
                case LogicOperator.Equals:
                    return AlmostEqual(left, right);
                case LogicOperator.NotEquals:
                    return !AlmostEqual(left, right);
                case LogicOperator.GreaterThan:
                    return left > right;
                case LogicOperator.GreaterThanOrEqual:
                    return left > right || AlmostEqual(left, right);
                case LogicOperator.LessThan:
                    return left < right;
                case LogicOperator.LessThanOrEqual:
                    return left < right || AlmostEqual(left, right);
                default:
                    return false;
            }
        }

        private static bool AlmostEqual(double left, double right, double tolerance = 1e-3)
        {
            return Math.Abs(left - right) <= tolerance;
        }

        private bool TryParseDouble(string input, out double value)
        {
            if (double.TryParse(input, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            return double.TryParse(input, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out value);
        }

        private static double GetConfidenceValue(ObjectInfo obj)
        {
            double result = obj.confidence;
            if (Math.Abs(result) < double.Epsilon)
            {
                result = obj.classifyScore;
            }
            return result;
        }

        private string FormatGroupText(RuleGroupNode group)
        {
            string descriptor = group.Operator == LogicalOperatorType.And ? "ALL" : "ANY";
            return $"Fail when {descriptor} child rules are TRUE";
        }

        private string FormatRuleText(RuleConditionNode rule)
        {
            string fieldLabel = GetFieldLabel(rule.Field);
            string operatorLabel = GetOperatorLabel(rule.Operator);
            string value = string.IsNullOrWhiteSpace(rule.Value) ? "(value)" : rule.Value;
            return $"Fail if {fieldLabel} {operatorLabel} {value}";
        }

        private string GetFieldLabel(LogicField field)
        {
            switch (field)
            {
                case LogicField.ClassName:
                    return "Class";
                case LogicField.Confidence:
                    return "Confidence";
                case LogicField.Area:
                    return "Area";
                case LogicField.Count:
                    return "Count";
                default:
                    return field.ToString();
            }
        }

        private string GetOperatorLabel(LogicOperator op)
        {
            switch (op)
            {
                case LogicOperator.Equals:
                    return "is";
                case LogicOperator.NotEquals:
                    return "is not";
                case LogicOperator.GreaterThan:
                    return ">";
                case LogicOperator.GreaterThanOrEqual:
                    return ">=";
                case LogicOperator.LessThan:
                    return "<";
                case LogicOperator.LessThanOrEqual:
                    return "<=";
                case LogicOperator.Contains:
                    return "contains";
                case LogicOperator.StartsWith:
                    return "starts with";
                case LogicOperator.EndsWith:
                    return "ends with";
                default:
                    return op.ToString();
            }
        }

        private LogicOperator UpdateOperatorOptions(LogicField field, LogicOperator? preferred = null)
        {
            List<ComboOption<LogicOperator>> options = GetOperatorOptions(field);
            List<ComboOption<LogicOperator>> dataSource = options.Select(option => new ComboOption<LogicOperator>(option.Display, option.Value)).ToList();

            bool previousSuppress = suppressLogicEvents;
            suppressLogicEvents = true;

            CB_Operator.DisplayMember = "Display";
            CB_Operator.ValueMember = "Value";
            CB_Operator.DataSource = dataSource;

            LogicOperator selected = preferred.HasValue && dataSource.Any(o => o.Value.Equals(preferred.Value))
                ? preferred.Value
                : dataSource.First().Value;
            CB_Operator.SelectedValue = selected;

            suppressLogicEvents = previousSuppress;
            return selected;
        }

        private void UpdateValueSuggestions(LogicField field, string preferredValue = null)
        {
            bool previousSuppress = suppressLogicEvents;
            suppressLogicEvents = true;

            List<ObjectInfo> defectDetections = CollectCurrentDefectDetections();
            List<string> suggestions = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddSuggestion(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return;
                }

                string trimmed = value.Trim();
                if (seen.Add(trimmed))
                {
                    suggestions.Add(trimmed);
                }
            }

            switch (field)
            {
                case LogicField.ClassName:
                    if (classNameList != null)
                    {
                        foreach (string name in classNameList)
                        {
                            AddSuggestion(name);
                        }
                    }

                    foreach (ObjectInfo detection in defectDetections)
                    {
                        if (!string.IsNullOrWhiteSpace(detection?.name))
                        {
                            AddSuggestion(detection.name);
                        }
                    }
                    break;

                case LogicField.Count:
                    AddSuggestion(defectDetections.Count.ToString(CultureInfo.InvariantCulture));
                    if (currentImage?.FrontInspections != null)
                    {
                        AddSuggestion(currentImage.FrontInspections.Count.ToString(CultureInfo.InvariantCulture));
                    }
                    break;

                case LogicField.Confidence:
                    foreach (double confidence in defectDetections
                        .Select(GetConfidenceValue)
                        .Where(v => !double.IsNaN(v)))
                    {
                        AddSuggestion(confidence.ToString("0.###", CultureInfo.InvariantCulture));
                    }
                    break;

                case LogicField.Area:
                    foreach (double area in defectDetections
                        .Select(obj => (double)(Math.Abs(obj.DisplayRec.Width) * Math.Abs(obj.DisplayRec.Height)))
                        .Where(v => v > 0))
                    {
                        AddSuggestion(area.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                    break;
            }

            if (field == LogicField.Count && suggestions.Count == 0)
            {
                AddSuggestion("0");
            }

            suggestions = suggestions.Take(25).ToList();

            CB_Value.BeginUpdate();
            CB_Value.Items.Clear();
            AutoCompleteStringCollection autoSource = new AutoCompleteStringCollection();
            foreach (string suggestion in suggestions)
            {
                CB_Value.Items.Add(suggestion);
                autoSource.Add(suggestion);
            }
            CB_Value.AutoCompleteCustomSource = autoSource;

            if (!string.IsNullOrWhiteSpace(preferredValue))
            {
                CB_Value.Text = preferredValue;
            }
            else if (suggestions.Count > 0)
            {
                CB_Value.Text = suggestions[0];
            }
            else
            {
                CB_Value.Text = string.Empty;
            }

            CB_Value.EndUpdate();

            suppressLogicEvents = previousSuppress;
        }

        private enum LogicField
        {
            ClassName,
            Confidence,
            Area,
            Count
        }

        private enum LogicOperator
        {
            Equals,
            NotEquals,
            GreaterThan,
            GreaterThanOrEqual,
            LessThan,
            LessThanOrEqual,
            Contains,
            StartsWith,
            EndsWith
        }

        private enum LogicalOperatorType
        {
            Or,
            And
        }

        private abstract class LogicNodeBase
        {
            public Guid Id { get; } = Guid.NewGuid();
        }

        private class RuleGroupNode : LogicNodeBase
        {
            public LogicalOperatorType Operator { get; set; } = LogicalOperatorType.Or;
            public List<LogicNodeBase> Children { get; } = new List<LogicNodeBase>();
        }

        private class RuleConditionNode : LogicNodeBase
        {
            public LogicField Field { get; set; } = LogicField.ClassName;
            public LogicOperator Operator { get; set; } = LogicOperator.Equals;
            public string Value { get; set; } = string.Empty;
        }

        private sealed class ModelContext : IDisposable
        {
            public ModelContext(string name, SolVision.TaskProcess process)
            {
                Name = name;
                Process = process;
            }

            public string Name { get; }
            public SolVision.TaskProcess Process { get; }
            public string LoadedPath { get; set; } = string.Empty;
            public bool IsLoaded { get; set; }
            public string LastError { get; set; }
            public TextBox PathDisplay { get; set; }
            public Label StatusDisplay { get; set; }
            public Button BrowseButton { get; set; }
            public Button LoadButton { get; set; }

            public void Dispose()
            {
                Process?.Close();
                Process?.CloseFromDLL();
            }
        }

        private enum CameraRole
        {
            Top,
            Front
        }

        private sealed class CameraDeviceInfo
        {
            public CameraDeviceInfo(
                int index,
                string name,
                string serial,
                string vendor,
                IMVDefine.IMV_ECameraType cameraType,
                IMVDefine.IMV_EInterfaceType interfaceType)
            {
                Index = index;
                Name = name ?? string.Empty;
                SerialNumber = serial ?? string.Empty;
                VendorName = vendor ?? string.Empty;
                CameraType = cameraType;
                InterfaceType = interfaceType;
            }

            public int Index { get; }
            public string Name { get; }
            public string SerialNumber { get; }
            public string VendorName { get; }
            public IMVDefine.IMV_ECameraType CameraType { get; }
            public IMVDefine.IMV_EInterfaceType InterfaceType { get; }

            public string DisplayName
            {
                get
                {
                    string vendorPrefix = string.IsNullOrWhiteSpace(VendorName) ? string.Empty : VendorName.Trim();
                    string cameraName = string.IsNullOrWhiteSpace(Name) ? "Camera" : Name.Trim();
                    string interfaceLabel = DescribeInterface(InterfaceType);

                    if (!string.IsNullOrEmpty(vendorPrefix) &&
                        !cameraName.StartsWith(vendorPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cameraName = $"{vendorPrefix} {cameraName}";
                    }

                    string serialPart = string.IsNullOrWhiteSpace(SerialNumber) ? string.Empty : $" [{SerialNumber}]";
                    return $"{cameraName}{serialPart} ({interfaceLabel})";
                }
            }

            public override string ToString() => DisplayName;
        }

        private sealed class CameraContext : IDisposable
        {
            public CameraContext(CameraRole role)
            {
                Role = role;
            }

            public CameraRole Role { get; }
            public ComboBox Selector { get; set; }
            public Button ConnectButton { get; set; }
            public Button CaptureButton { get; set; }
            public Label StatusLabel { get; set; }
            public Label DetailLabel { get; set; }
            public MyCamera Camera { get; private set; }
            public CameraDeviceInfo ConnectedDevice { get; private set; }
            public IntPtr ConversionBuffer { get; set; }
            public int ConversionBufferSize { get; set; }

            public string RoleName => Role == CameraRole.Top ? "Top" : "Front";
            public bool IsConnected => Camera != null && ConnectedDevice != null;

            public void SetConnectedDevice(CameraDeviceInfo device, MyCamera camera)
            {
                ConnectedDevice = device;
                Camera = camera;
                if (DetailLabel != null && device != null)
                {
                    DetailLabel.Text = $"{device.DisplayName}\nCapture a frame to refresh the preview.";
                }
            }

            public void Disconnect()
            {
                if (Camera != null)
                {
                    try
                    {
                        if (Camera.IMV_IsGrabbing())
                        {
                            Camera.IMV_StopGrabbing();
                        }
                    }
                    catch
                    {
                        // ignore teardown errors
                    }

                    try
                    {
                        if (Camera.IMV_IsOpen())
                        {
                            Camera.IMV_Close();
                        }
                    }
                    catch
                    {
                        // ignore teardown errors
                    }

                    try
                    {
                        Camera.IMV_DestroyHandle();
                    }
                    catch
                    {
                        // ignore teardown errors
                    }
                }

                Camera = null;
                ConnectedDevice = null;
                ClearPreview();
                if (DetailLabel != null)
                {
                    DetailLabel.Text = $"Choose which device should act as the {RoleName.ToLowerInvariant()} camera.";
                }
            }

            public void ClearPreview()
            {
                // no inline preview panel in camera card, nothing to clear
            }

            public void Dispose()
            {
                Disconnect();
                if (ConversionBuffer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(ConversionBuffer);
                    ConversionBuffer = IntPtr.Zero;
                }
                ConversionBufferSize = 0;
            }
        }

        private class ComboOption<T>
        {
            public ComboOption(string display, T value)
            {
                Display = display;
                Value = value;
            }

            public string Display { get; }
            public T Value { get; }

            public override string ToString() => Display;
        }
        private readonly struct TurntableHomeResult

        {

            public TurntableHomeResult(bool success, string message, double? offsetDegrees)

            {

                Success = success;

                Message = message ?? string.Empty;

                OffsetDegrees = offsetDegrees;

            }



            public bool Success { get; }

            public string Message { get; }

            public double? OffsetDegrees { get; }



            public static TurntableHomeResult Failed(string message)

            {

                return new TurntableHomeResult(false, message, null);

            }

        }

        private sealed class TurntableController : IDisposable
        {
            private const int CommandGuardMilliseconds = 100;

            private SerialPort serialPort;

            private readonly object commandLock = new object();

            private readonly StringBuilder receiveBuffer = new StringBuilder();

            private DateTime lastCommandUtc = DateTime.MinValue;



            public bool IsConnected => serialPort?.IsOpen == true;

            public bool IsHomed { get; private set; }

            public double? LastOffsetAngle { get; private set; }

            public string PortName => serialPort?.PortName;



            public event Action<string> MessageReceived;



            public void Connect(string portName, int baudRate = 115200)

            {

                if (string.IsNullOrWhiteSpace(portName))

                {

                    throw new ArgumentException("Port name cannot be empty.", nameof(portName));

                }



                Disconnect();



                SerialPort port = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)

                {

                    Encoding = Encoding.ASCII,

                    NewLine = ";",

                    ReadTimeout = 200,

                    WriteTimeout = 200,

                    Handshake = Handshake.None

                };



                try

                {

                    port.DataReceived += SerialPortDataReceived;

                    port.Open();

                    port.DiscardInBuffer();

                    port.DiscardOutBuffer();

                    serialPort = port;

                }

                catch

                {

                    port.DataReceived -= SerialPortDataReceived;

                    port.Dispose();

                    throw;

                }



                IsHomed = false;

                LastOffsetAngle = null;

                lastCommandUtc = DateTime.MinValue;



                OnMessage($"INFO Connected to {portName}");

            }



            public void Disconnect()

            {

                SerialPort port = serialPort;

                if (port == null)

                {

                    return;

                }



                serialPort = null;



                try

                {

                    port.DataReceived -= SerialPortDataReceived;

                    if (port.IsOpen)

                    {

                        port.Close();

                    }

                }

                finally

                {

                    string closedPort = port.PortName;

                    port.Dispose();

                    IsHomed = false;

                    LastOffsetAngle = null;

                    OnMessage($"INFO Disconnected from {closedPort}");

                }

            }



            public async Task<TurntableHomeResult> HomeAsync(CancellationToken cancellationToken)
            {
                if (!IsConnected)
                {
                    return TurntableHomeResult.Failed("Turntable not connected.");
                }

                IsHomed = false;
                LastOffsetAngle = null;

                try
                {
                    double offset = await RequestOffsetAsync(cancellationToken).ConfigureAwait(false);
                    LastOffsetAngle = offset;

                    double moveAngle = Math.Abs(offset);
                    if (IsAngleAlreadyZero(moveAngle))
                    {
                        IsHomed = true;
                        LastOffsetAngle = 0;
                        return new TurntableHomeResult(true, "Homing complete (already at zero).", offset);
                    }

                    string summary = await ExecuteHomingRotationAsync(offset, cancellationToken).ConfigureAwait(false);
                    IsHomed = true;
                    LastOffsetAngle = 0;
                    return new TurntableHomeResult(true, summary, offset);
                }
                catch (OperationCanceledException)
                {
                    IsHomed = false;
                    return TurntableHomeResult.Failed("Homing cancelled.");
                }
                catch (Exception ex)
                {
                    IsHomed = false;
                    return TurntableHomeResult.Failed(ex.Message);
                }
            }



        public async Task<string> MoveRelativeAsync(double angleDegrees, CancellationToken cancellationToken)
        {
            if (!IsConnected)
            {
                throw new InvalidOperationException("Turntable is not connected.");
            }

            double magnitude = Math.Abs(angleDegrees);
            if (magnitude < 1e-3)
            {
                return "Rotation skipped (below threshold).";
            }

            int direction = angleDegrees >= 0 ? 0 : 1; // 0 = CW, 1 = CCW per vendor protocol
            string directionLabel = direction == 0 ? "CW" : "CCW";

            var ackSource = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var eventSource = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

            void Handler(string message)
            {
                string normalized = NormalizeMessage(message);
                if (string.IsNullOrEmpty(normalized))
                {
                    return;
                }

                if (!ackSource.Task.IsCompleted)
                {
                    if (normalized.StartsWith("CR+ERR", StringComparison.OrdinalIgnoreCase))
                    {
                        var error = new InvalidOperationException(normalized);
                        ackSource.TrySetException(error);
                        eventSource.TrySetException(error);
                        return;
                    }

                    if (normalized.StartsWith("CR+OK", StringComparison.OrdinalIgnoreCase))
                    {
                        ackSource.TrySetResult(normalized);
                        return;
                    }
                }

                if (normalized.StartsWith("CR+EVENT=TB_END", StringComparison.OrdinalIgnoreCase))
                {
                    eventSource.TrySetResult(normalized);
                }
            }

            MessageReceived += Handler;

            using (cancellationToken.Register(() =>
            {
                ackSource.TrySetCanceled(cancellationToken);
                eventSource.TrySetCanceled(cancellationToken);
            }, useSynchronizationContext: false))
            {
                try
                {
                    string command = $"CT+START({direction},1,0,{magnitude.ToString("F4", CultureInfo.InvariantCulture)},0,1);";
                    SendCommand(command);

                    await ackSource.Task.ConfigureAwait(false);
                    await eventSource.Task.ConfigureAwait(false);

                    return $"Rotated {magnitude:0.00} deg {directionLabel}.";
                }
                finally
                {
                    MessageReceived -= Handler;
                }
            }
        }

        public void SendCommand(string command)

            {

                if (!IsConnected)

                {

                    throw new InvalidOperationException("Turntable is not connected.");

                }



                if (string.IsNullOrWhiteSpace(command))

                {

                    throw new ArgumentException("Command cannot be empty.", nameof(command));

                }



                lock (commandLock)

                {

                    double elapsed = (DateTime.UtcNow - lastCommandUtc).TotalMilliseconds;

                    if (elapsed < CommandGuardMilliseconds)

                    {

                        Thread.Sleep(CommandGuardMilliseconds - (int)elapsed);

                    }



                    SerialPort port = serialPort;

                    if (port == null)

                    {

                        throw new InvalidOperationException("Turntable connection is closed.");

                    }



                    string payload = command.EndsWith(";", StringComparison.Ordinal) ? command : command + ";";

                    port.Write(payload);

                    lastCommandUtc = DateTime.UtcNow;

                }

            }



            private async Task<double> RequestOffsetAsync(CancellationToken cancellationToken)

            {

                var completion = new TaskCompletionSource<double>(TaskCreationOptions.RunContinuationsAsynchronously);



                void Handler(string message)

                {

                    string normalized = NormalizeMessage(message);

                    if (string.IsNullOrEmpty(normalized))

                    {

                        return;

                    }



                    if (normalized.StartsWith("CR+ERR", StringComparison.OrdinalIgnoreCase))

                    {

                        completion.TrySetException(new InvalidOperationException(normalized));

                        return;

                    }



                    if (normalized.IndexOf("OffsetAngle", StringComparison.OrdinalIgnoreCase) >= 0)

                    {

                        double? offset = TryParseOffset(normalized);

                        if (offset.HasValue)

                        {

                            completion.TrySetResult(offset.Value);

                        }

                        else

                        {

                            completion.TrySetException(new InvalidOperationException("Unable to parse OffsetAngle response."));

                        }

                    }

                }



                MessageReceived += Handler;



                using (cancellationToken.Register(() => completion.TrySetCanceled(cancellationToken), useSynchronizationContext: false))

                {

                    try

                    {

                        SendCommand("CT+GETOFFSETANGLE();");

                        return await completion.Task.ConfigureAwait(false);

                    }

                    finally

                    {

                        MessageReceived -= Handler;

                    }

                }

            }



            private async Task<string> ExecuteHomingRotationAsync(double offset, CancellationToken cancellationToken)

            {

                double moveAngle = Math.Abs(offset);

                int direction = offset > 0 ? 1 : 0;

                string directionLabel = direction == 1 ? "CCW" : "CW";



                var ackSource = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

                var eventSource = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);



                void Handler(string message)

                {

                    string normalized = NormalizeMessage(message);

                    if (string.IsNullOrEmpty(normalized))

                    {

                        return;

                    }



                    if (!ackSource.Task.IsCompleted)

                    {

                        if (normalized.StartsWith("CR+ERR", StringComparison.OrdinalIgnoreCase))

                        {

                            ackSource.TrySetException(new InvalidOperationException(normalized));

                            eventSource.TrySetException(new InvalidOperationException(normalized));

                            return;

                        }



                        if (normalized.StartsWith("CR+OK", StringComparison.OrdinalIgnoreCase))

                        {

                            ackSource.TrySetResult(normalized);

                            return;

                        }

                    }



                    if (normalized.StartsWith("CR+EVENT=TB_END", StringComparison.OrdinalIgnoreCase))

                    {

                        eventSource.TrySetResult(normalized);

                    }

                    else if (normalized.StartsWith("CR+ERR", StringComparison.OrdinalIgnoreCase))

                    {

                        eventSource.TrySetException(new InvalidOperationException(normalized));

                    }

                }



                MessageReceived += Handler;



                using (cancellationToken.Register(() =>

                {

                    ackSource.TrySetCanceled(cancellationToken);

                    eventSource.TrySetCanceled(cancellationToken);

                }, useSynchronizationContext: false))

                {

                    try

                    {

                        string command = $"CT+START({direction},1,0,{moveAngle.ToString("F4", CultureInfo.InvariantCulture)},0,1);";

                        SendCommand(command);



                        await ackSource.Task.ConfigureAwait(false);

                        string evt = await eventSource.Task.ConfigureAwait(false);

                        if (!evt.StartsWith("CR+EVENT=TB_END", StringComparison.OrdinalIgnoreCase))

                        {

                            throw new InvalidOperationException($"Unexpected response: {evt}");

                        }



                        return $"Homing complete. Rotated {moveAngle:0.00} deg {directionLabel} to zero.";

                    }

                    finally

                    {

                        MessageReceived -= Handler;

                    }

                }

            }



            private static bool IsAngleAlreadyZero(double moveAngle)

            {

                return moveAngle < 1e-3 || Math.Abs(moveAngle - 360.0) < 1e-3;

            }



            public static string NormalizeMessage(string message)

            {

                if (string.IsNullOrWhiteSpace(message))

                {

                    return string.Empty;

                }



                string normalized = message.Trim();

                normalized = normalized.TrimStart('#').Trim();

                if (normalized.EndsWith(";", StringComparison.Ordinal))

                {

                    normalized = normalized.Substring(0, normalized.Length - 1);

                }



                return normalized.Trim();

            }



            public void Dispose()

            {

                Disconnect();

            }



            private static double? TryParseOffset(string message)
            {
                int index = message.IndexOf("OffsetAngle=", StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                {
                    return null;
                }

                string valueText = message.Substring(index + "OffsetAngle=".Length).Trim();
                valueText = valueText.TrimEnd(';');

                Match match = Regex.Match(valueText, @"-?\d+(?:\.\d+)?", RegexOptions.CultureInvariant);
                if (!match.Success)
                {
                    return null;
                }

                if (double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double offset))
                {
                    return offset;
                }

                return null;
            }



            private void SerialPortDataReceived(object sender, SerialDataReceivedEventArgs e)

            {

                SerialPort port = serialPort;

                if (port == null)

                {

                    return;

                }



                try

                {

                    string chunk = port.ReadExisting();

                    if (string.IsNullOrEmpty(chunk))

                    {

                        return;

                    }



                    lock (receiveBuffer)

                    {

                        receiveBuffer.Append(chunk);

                        while (true)

                        {

                            string bufferText = receiveBuffer.ToString();

                            int delimiter = bufferText.IndexOf(';');

                            if (delimiter < 0)

                            {

                                break;

                            }



                            string message = bufferText.Substring(0, delimiter + 1);

                            receiveBuffer.Remove(0, delimiter + 1);

                            OnMessage(message);

                        }

                    }

                }

                catch (Exception ex)

                {

                    OnMessage($"ERR Serial read error: {ex.Message}");

                }

            }



            private void OnMessage(string message)

            {

                MessageReceived?.Invoke(message);

            }

        }






            
    }



internal sealed class AttachmentPointInfo
{
    public AttachmentPointInfo(ObjectInfo source, PointF center, double angleDegrees, int sequence)
    {
        Source = source;
        Center = center;
        AngleDegrees = angleDegrees;
        Sequence = sequence;
    }

    public ObjectInfo Source { get; }
    public PointF Center { get; }
    public double AngleDegrees { get; }
    public int Sequence { get; }
    public string CapturedImagePath { get; set; }
    public FrontInspectionResult FrontInspection { get; set; }
}

internal sealed class AttachmentOverlayResult
{
    public AttachmentOverlayResult(Bitmap overlayBitmap, List<AttachmentPointInfo> points, PointF center)
    {
        OverlayBitmap = overlayBitmap;
        Points = points ?? new List<AttachmentPointInfo>();
        Center = center;
    }

    public Bitmap OverlayBitmap { get; }
    public List<AttachmentPointInfo> Points { get; }
    public PointF Center { get; }
}

internal sealed class FrontCaptureTicket
{
    public FrontCaptureTicket(int sequence, string imagePath, double angleDegrees, DateTime capturedAt)
    {
        Sequence = sequence;
        ImagePath = imagePath;
        AngleDegrees = angleDegrees;
        CapturedAt = capturedAt;
    }

    public int Sequence { get; }
    public string ImagePath { get; }
    public double AngleDegrees { get; }
    public DateTime CapturedAt { get; }
}

internal sealed class FrontInspectionResult : IDisposable
{
    public FrontInspectionResult(int sequence, string imagePath, double angleDegrees, DateTime capturedAt, List<ObjectInfo> detections, Bitmap overlayImage, Bitmap rawImage)
    {
        Sequence = sequence;
        ImagePath = imagePath;
        AngleDegrees = angleDegrees;
        CapturedAt = capturedAt;
        Detections = detections ?? new List<ObjectInfo>();
        OverlayImage = overlayImage;
        RawImage = rawImage;
    }

    public int Sequence { get; }
    public string ImagePath { get; }
    public double AngleDegrees { get; }
    public DateTime CapturedAt { get; }
    public List<ObjectInfo> Detections { get; }
    public Bitmap OverlayImage { get; private set; }
    public Bitmap RawImage { get; private set; }
    public bool HasDefects => Detections != null && Detections.Count > 0;

    public Dictionary<string, int> GetClassHistogram()
    {
        Dictionary<string, int> histogram = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (Detections == null)
        {
            return histogram;
        }

        foreach (ObjectInfo obj in Detections)
        {
            string key = string.IsNullOrWhiteSpace(obj?.name) ? "(unknown)" : obj.name.Trim();
            if (histogram.ContainsKey(key))
            {
                histogram[key]++;
            }
            else
            {
                histogram[key] = 1;
            }
        }
        return histogram;
    }

    public string BuildSummary()
    {
        Dictionary<string, int> histogram = GetClassHistogram();
        if (histogram.Count == 0)
        {
            return "No defects";
        }

        return string.Join(", ", histogram.Select(pair => $"{pair.Key} ×{pair.Value}"));
    }

    public void UpdateOverlay(Bitmap overlayImage, Bitmap rawImage)
    {
        OverlayImage?.Dispose();
        RawImage?.Dispose();
        OverlayImage = overlayImage;
        RawImage = rawImage;
    }

    public void Dispose()
    {
        OverlayImage?.Dispose();
        OverlayImage = null;
        RawImage?.Dispose();
        RawImage = null;
        Detections?.Clear();
    }
}

internal class DetectImg : IDisposable

    {
        internal Mat OriImg;
        internal string OriImgPath;
        internal List<ObjectInfo> ObjList = new List<ObjectInfo>();
        internal List<AttachmentPointInfo> AttachmentPoints = new List<AttachmentPointInfo>();
        internal PointF AttachmentCenter;
        internal bool HasResults;
        internal List<FrontInspectionResult> FrontInspections = new List<FrontInspectionResult>();
        internal bool FrontInspectionComplete;

        public void Dispose()
        {
            OriImg?.Dispose();
            OriImg = null;
            ObjList?.Clear();
            HasResults = false;
            if (AttachmentPoints != null)
            {
                foreach (AttachmentPointInfo point in AttachmentPoints)
                {
                    if (point != null)
                    {
                        point.FrontInspection = null;
                    }
                }
                AttachmentPoints.Clear();
            }
            AttachmentCenter = new PointF();
            if (FrontInspections != null)
            {
                foreach (FrontInspectionResult inspection in FrontInspections)
                {
                    inspection?.Dispose();
                }
                FrontInspections.Clear();
            }
            FrontInspectionComplete = false;
        }
    }
}










