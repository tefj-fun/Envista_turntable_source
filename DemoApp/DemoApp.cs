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
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Drawing.Text;
using MVSDK_Net;
using Newtonsoft.Json;

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
        private readonly Color statusPassBackground = Color.FromArgb(204, 232, 208);
        private readonly Color statusPassForeground = Color.FromArgb(20, 80, 30);
        private readonly Color statusFailBackground = Color.FromArgb(255, 210, 214);
        private readonly Color statusFailForeground = Color.FromArgb(178, 34, 34);
        private readonly Color statusNeutralBackground = Color.FromArgb(245, 247, 250);
        private readonly Color statusNeutralForeground = Color.FromArgb(94, 102, 112);
        private const string StatusPassNoRulesMessage = "Result: PASS (no rules defined)";
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
        private GroupBox groupClassRules;
        private FlowLayoutPanel flowClassRules;
        private readonly Dictionary<string, ClassRuleCard> classRuleCards = new Dictionary<string, ClassRuleCard>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Image> glyphImageCache = new Dictionary<string, Image>();

        private ModelContext attachmentContext;
        private ModelContext frontAttachmentContext;
        private ModelContext defectContext;

        private TableLayoutPanel initLayout;
        private GroupBox groupInitAttachment;
        private GroupBox groupInitFrontAttachment;
        private GroupBox groupInitDefect;
        private TextBox TB_InitAttachmentPath;
        private Button BT_InitAttachmentBrowse;
        private Button BT_InitAttachmentLoad;
        private Label LBL_InitAttachmentStatus;
        private TextBox TB_InitFrontAttachmentPath;
        private Button BT_InitFrontAttachmentBrowse;
        private Button BT_InitFrontAttachmentLoad;
        private Label LBL_InitFrontAttachmentStatus;
        private TextBox TB_InitDefectPath;
        private Button BT_InitDefectBrowse;
        private Button BT_InitDefectLoad;
        private Label LBL_InitDefectStatus;
        private Label LBL_InitSummary;

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
        private ComboBox CB_TurntablePort;
        private Button BT_TurntableRefresh;
        private Button BT_TurntableConnect;
        private Button BT_TurntableHome;
        private Label LBL_TurntableStatus;
        private Label LBL_StepModelsStatus;
        private Label LBL_StepCamerasStatus;
        private Label LBL_StepTurntableStatus;
        private Button BT_InitBeginWorkflow;
        private Button BT_OpenInitWizard;
        private Label LBL_InitPrompt;
        private Form activeInitDialog;
        private Label LBL_InitSummaryModal;
        private GroupBox groupWorkflowInit;
        private bool initializationPromptShown;
        private InitializationSettings initSettings;
        private DefectPolicy defectPolicy;
        private readonly string settingsRootDirectory;
        private readonly string initSettingsPath;
        private readonly string defectPolicyPath;
        private bool useRecordedRun;
        private string recordedRunPath;
        private CheckBox CHK_UseRecordedRun;
        private TextBox TB_RecordedRunPath;
        private Button BT_SelectRecordedRun;
        private RunSession currentRunSession;
        private bool isUpdatingRecordedRunUI;
        private FlowLayoutPanel recordedRunPanel;

        private CameraContext topCameraContext;
        private CameraContext frontCameraContext;
        private readonly List<CameraDeviceInfo> cameraDeviceCache = new List<CameraDeviceInfo>();
        private CancellationTokenSource captureSequenceCts;
        private bool showFrontOverlay = true;
        private int selectedFrontSequence = -1;
        private Stopwatch cycleTimer;
        private Label LBL_CycleTime;
        private string currentPartID;
        private TextBox TB_PartID;
        private Label LBL_PartID;
        private static readonly object consoleCaptureLock = new object();
        private static bool consoleCaptureInitialized;
        private static ConsoleRedirectWriter consoleOutInterceptor;
        private static ConsoleRedirectWriter consoleErrorInterceptor;
        private static TextWriter consoleOutOriginal;
        private static TextWriter consoleErrorOriginal;
        private readonly object sdkTaskLock = new object();
        private readonly Queue<string> sdkTaskQueue = new Queue<string>();
        private readonly object consoleQueueLock = new object();
        private readonly Queue<Tuple<LogStatus, string>> consolePendingQueue = new Queue<Tuple<LogStatus, string>>();
        private bool responsiveLayoutInitialized;
        private bool usingCompactLayout;
        private const int CompactLayoutWidthThreshold = 1400;
        private const float CompactLayoutDpiThreshold = 1.5f;
        private const string IconFontFamily = "Segoe MDL2 Assets";
        private float CurrentDpiScale => DeviceDpi / 96f;

        public DemoApp()
        {
            InitializeComponent();
            EnsureConsoleCapture();
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            settingsRootDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SolVisionDemoApp");
            try
            {
                Directory.CreateDirectory(settingsRootDirectory);
            }
            catch
            {
                // Safe to ignore; individual save operations handle failures.
            }

            initSettingsPath = Path.Combine(settingsRootDirectory, "init-settings.ini");
            defectPolicyPath = Path.Combine(settingsRootDirectory, "defect-policy.json");
            initSettings = InitializationSettings.Load(initSettingsPath);
            defectPolicy = DefectPolicy.Load(defectPolicyPath);
            if (defectPolicy == null)
            {
                defectPolicy = DefectPolicy.Create();
            }
            useRecordedRun = initSettings?.UseRecordedRun ?? false;
            recordedRunPath = initSettings?.RecordedRunPath;
            classNameList = new List<string>();
            topCameraContext = new CameraContext(CameraRole.Top);
            frontCameraContext = new CameraContext(CameraRole.Front);
            turntableController = new TurntableController();
            turntableController.MessageReceived += TurntableController_MessageReceived;
            BuildLogicBuilderUI();
            BuildWorkflowInitializationCard();
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
            InitializeResponsiveLayout();
            InitialSolVision();
            LoadInitializationSettings();
            ApplyTheme();
            InitializeLogicBuilder();
            InitializeCycleTimeLabel();
            Shown += DemoApp_Shown;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            DrainPendingConsoleQueue();
        }

        private void DrainPendingConsoleQueue()
        {
            Queue<Tuple<LogStatus, string>> pending = null;
            lock (consoleQueueLock)
            {
                if (consolePendingQueue.Count > 0)
                {
                    pending = new Queue<Tuple<LogStatus, string>>(consolePendingQueue);
                    consolePendingQueue.Clear();
                }
            }

            if (pending == null || pending.Count == 0)
            {
                return;
            }

            while (pending.Count > 0)
            {
                Tuple<LogStatus, string> entry = pending.Dequeue();
                ProcessSdkLogLine(entry.Item2, entry.Item1);
            }
        }

        private Font ScaleFont(string family, float size, FontStyle style = FontStyle.Regular, GraphicsUnit unit = GraphicsUnit.Point)
        {
            float scaledSize = Math.Max(1f, size * CurrentDpiScale);
            return new Font(family, scaledSize, style, unit);
        }

        private int ScaleSize(int value)
        {
            return Math.Max(1, (int)Math.Round(value * CurrentDpiScale));
        }

        private void DemoApp_Shown(object sender, EventArgs e)
        {
            if (initializationPromptShown)
            {
                return;
            }

            initializationPromptShown = true;
            EnsureInitLayout();
            UpdateInitSummary();

            if (!IsInitializationComplete())
            {
                ShowInitializationWizard("Complete the initialization wizard before beginning the workflow.");
            }
        }

        private void EnsureConsoleCapture()
        {
            lock (consoleCaptureLock)
            {
                if (consoleCaptureInitialized)
                {
                    return;
                }

                consoleOutOriginal = Console.Out;
                consoleErrorOriginal = Console.Error;

                consoleOutInterceptor = new ConsoleRedirectWriter(consoleOutOriginal, HandleSdkConsoleMessage, LogStatus.Info);
                consoleErrorInterceptor = new ConsoleRedirectWriter(consoleErrorOriginal, HandleSdkConsoleMessage, LogStatus.Error);

                Console.SetOut(consoleOutInterceptor);
                Console.SetError(consoleErrorInterceptor);

                consoleCaptureInitialized = true;
            }
        }

        private void ReleaseConsoleCapture()
        {
            lock (consoleCaptureLock)
            {
                if (!consoleCaptureInitialized)
                {
                    return;
                }

                consoleOutInterceptor?.FlushPending();
                consoleErrorInterceptor?.FlushPending();

                if (consoleOutOriginal != null)
                {
                    Console.SetOut(consoleOutOriginal);
                }

                if (consoleErrorOriginal != null)
                {
                    Console.SetError(consoleErrorOriginal);
                }

                consoleOutInterceptor?.Dispose();
                consoleErrorInterceptor?.Dispose();
                consoleOutInterceptor = null;
                consoleErrorInterceptor = null;
                consoleCaptureInitialized = false;
            }
        }

        private void HandleSdkConsoleMessage(LogStatus status, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string trimmed = message.Trim();
            if (!IsHandleCreated)
            {
                lock (consoleQueueLock)
                {
                    consolePendingQueue.Enqueue(Tuple.Create(status, trimmed));
                }
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action(() => HandleSdkConsoleMessage(status, trimmed)));
                }
                catch (InvalidOperationException)
                {
                    lock (consoleQueueLock)
                    {
                        consolePendingQueue.Enqueue(Tuple.Create(status, trimmed));
                    }
                }
                return;
            }

            ProcessSdkLogLine(trimmed, status);
        }

        private void ProcessSdkLogLine(string message, LogStatus defaultStatus)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            if (message.StartsWith("TaskProcess:", StringComparison.OrdinalIgnoreCase))
            {
                string detail = ExtractAfter(message, "TaskProcess:").Trim();
                if (string.IsNullOrEmpty(detail))
                {
                    detail = "TaskProcess";
                }

                lock (sdkTaskLock)
                {
                    sdkTaskQueue.Enqueue(detail);
                }

                outToLog(message, LogStatus.Progress);
                return;
            }

            if (message.StartsWith("------task", StringComparison.OrdinalIgnoreCase))
            {
                string taskName = null;
                lock (sdkTaskLock)
                {
                    if (sdkTaskQueue.Count > 0)
                    {
                        taskName = sdkTaskQueue.Dequeue();
                    }
                }

                Match match = Regex.Match(message, @"^-{6}task\s+(?<index>\d+)\s*:\s*(?<duration>\d+)\s*ms", RegexOptions.IgnoreCase);
                string duration = match.Success ? match.Groups["duration"].Value : null;

                if (!string.IsNullOrEmpty(taskName))
                {
                    string completion = !string.IsNullOrEmpty(duration)
                        ? $"Task '{taskName}' completed in {duration} ms."
                        : $"Task '{taskName}' completed.";
                    outToLog(completion, LogStatus.Success);
                    return;
                }

                outToLog(message, LogStatus.Success);
                return;
            }

            outToLog(message, defaultStatus);
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

    SaveDefectPolicyToDisk();
    ReleaseConsoleCapture();

}



        public void InitialSolVision()
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application();
            }

            attachmentContext = CreateModelContext("Attachment");
            frontAttachmentContext = CreateModelContext("Front Attachment");
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

            if (frontAttachmentContext != null)
            {
                frontAttachmentContext.PathDisplay = TB_InitFrontAttachmentPath;
                frontAttachmentContext.StatusDisplay = LBL_InitFrontAttachmentStatus;
                frontAttachmentContext.BrowseButton = BT_InitFrontAttachmentBrowse;
                frontAttachmentContext.LoadButton = BT_InitFrontAttachmentLoad;
                UpdateModelStatus(frontAttachmentContext, "Not loaded.", statusNeutralBackground, statusNeutralForeground);
                ToggleInitButtons(frontAttachmentContext, true);
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
            UpdateInitSummary();
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
            if (initSettings != null)
            {
                if (context == attachmentContext)
                {
                    initSettings.AttachmentPath = filePath;
                }
                else if (context == frontAttachmentContext)
                {
                    initSettings.FrontAttachmentPath = filePath;
                }
                else if (context == defectContext)
                {
                    initSettings.DefectPath = filePath;
                }
                SaveInitializationSettings();
            }
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
                    if (initSettings != null)
                    {
                        initSettings.AttachmentPath = dlg.FileName;
                        SaveInitializationSettings();
                    }
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
                    if (initSettings != null)
                    {
                        initSettings.DefectPath = dlg.FileName;
                        SaveInitializationSettings();
                    }
                }
            }
        }

        private void BT_InitFrontAttachmentBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "SolVision Project (*.tsp)|*.tsp";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    TB_InitFrontAttachmentPath.Text = dlg.FileName;
                    ToggleInitButtons(frontAttachmentContext, true);
                    UpdateModelStatus(frontAttachmentContext, "Ready to load.", statusNeutralBackground, statusNeutralForeground);
                    if (initSettings != null)
                    {
                        initSettings.FrontAttachmentPath = dlg.FileName;
                        SaveInitializationSettings();
                    }
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

        private async void BT_InitFrontAttachmentLoad_Click(object sender, EventArgs e)
        {
            if (frontAttachmentContext == null)
            {
                outToLog("Front Attachment context is not ready.", LogStatus.Error);
                return;
            }

            string path = TB_InitFrontAttachmentPath.Text;
            if (!File.Exists(path))
            {
                UpdateModelStatus(frontAttachmentContext, "Project file not found.", statusFailBackground, statusFailForeground);
                outToLog($"[Front Attachment] Project file not found: {path}", LogStatus.Error);
                return;
            }

            await LoadProjectAsync(frontAttachmentContext, path, updateWorkflowPath: false);
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

        private void LoadInitializationSettings()
        {
            if (initSettings == null)
            {
                initSettings = new InitializationSettings();
            }

            EnsureInitLayout();

            if (attachmentContext != null && !string.IsNullOrWhiteSpace(initSettings.AttachmentPath))
            {
                if (File.Exists(initSettings.AttachmentPath))
                {
                    TB_InitAttachmentPath.Text = initSettings.AttachmentPath;
                    ToggleInitButtons(attachmentContext, true);
                    UpdateModelStatus(attachmentContext, "Ready to load (remembered).", statusNeutralBackground, statusNeutralForeground);
                }
                else
                {
                    outToLog($"[Settings] Attachment project not found: {initSettings.AttachmentPath}", LogStatus.Warning);
                }
            }

            if (frontAttachmentContext != null && !string.IsNullOrWhiteSpace(initSettings.FrontAttachmentPath))
            {
                if (File.Exists(initSettings.FrontAttachmentPath))
                {
                    TB_InitFrontAttachmentPath.Text = initSettings.FrontAttachmentPath;
                    ToggleInitButtons(frontAttachmentContext, true);
                    UpdateModelStatus(frontAttachmentContext, "Ready to load (remembered).", statusNeutralBackground, statusNeutralForeground);
                }
                else
                {
                    outToLog($"[Settings] Front Attachment project not found: {initSettings.FrontAttachmentPath}", LogStatus.Warning);
                }
            }

            if (defectContext != null && !string.IsNullOrWhiteSpace(initSettings.DefectPath))
            {
                if (File.Exists(initSettings.DefectPath))
                {
                    TB_InitDefectPath.Text = initSettings.DefectPath;
                    ToggleInitButtons(defectContext, true);
                    UpdateModelStatus(defectContext, "Ready to load (remembered).", statusNeutralBackground, statusNeutralForeground);
                }
                else
                {
                    outToLog($"[Settings] Defect project not found: {initSettings.DefectPath}", LogStatus.Warning);
                }
            }

            PopulateTurntablePorts();
            useRecordedRun = initSettings?.UseRecordedRun ?? false;
            recordedRunPath = initSettings?.RecordedRunPath;

            // Load last used Part ID
            if (!string.IsNullOrWhiteSpace(initSettings?.LastPartID))
            {
                currentPartID = initSettings.LastPartID;
                if (TB_PartID != null && !TB_PartID.IsDisposed)
                {
                    TB_PartID.Text = currentPartID;
                }
            }

            UpdateRecordedRunUiState();
            UpdateInitSummary();
            AdjustRecordedRunLayout();
        }

        private void SaveInitializationSettings()
        {
            if (initSettings == null)
            {
                initSettings = new InitializationSettings();
            }

            try
            {
                initSettings.Save(initSettingsPath);
            }
            catch (Exception ex)
            {
                outToLog($"[Settings] Failed to save initialization settings: {ex.Message}", LogStatus.Warning);
            }
        }

        private void SyncDefectPolicyWithClasses(IEnumerable<string> classes)
        {
            if (defectPolicy == null)
            {
                defectPolicy = DefectPolicy.Create();
            }

            if (classes == null)
            {
                return;
            }

            List<string> normalized = classes
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (normalized.Count == 0)
            {
                return;
            }

            bool changed = defectPolicy.SyncWithClasses(normalized);
            if (changed)
            {
                SaveDefectPolicyToDisk();
            }

            RefreshClassRuleCardsUI();
        }

        private void SaveDefectPolicyToDisk()
        {
            if (defectPolicy == null || string.IsNullOrWhiteSpace(defectPolicyPath))
            {
                return;
            }

            try
            {
                defectPolicy.Save(defectPolicyPath);
            }
            catch (Exception ex)
            {
                outToLog($"[Policy] Failed to save defect policy: {ex.Message}", LogStatus.Warning);
            }
        }

        private bool IsRecordedRunSelectionValid(out string message)
        {
            message = string.Empty;
            if (string.IsNullOrWhiteSpace(recordedRunPath))
            {
                message = "Select a saved run folder.";
                return false;
            }

            if (!Directory.Exists(recordedRunPath))
            {
                message = "Saved run folder does not exist.";
                return false;
            }

            if (!TryGetRecordedTopImagePath(recordedRunPath, out _))
            {
                message = "Saved run is missing a top image.";
                return false;
            }

            string frontFolder = Path.Combine(recordedRunPath, "Front");
            if (!Directory.Exists(frontFolder))
            {
                message = "Saved run is missing a Front folder.";
                return false;
            }

            if (Directory.GetFiles(frontFolder, "*.png").Length == 0)
            {
                message = "Saved run has no front images.";
                return false;
            }

            return true;
        }

        private void UpdateRecordedRunUiState()
        {
            bool valid = IsRecordedRunSelectionValid(out string validationMessage);
            isUpdatingRecordedRunUI = true;
            try
            {
                if (CHK_UseRecordedRun != null)
                {
                    CHK_UseRecordedRun.Checked = useRecordedRun;
                }

                if (TB_RecordedRunPath != null)
                {
                    TB_RecordedRunPath.Enabled = useRecordedRun;
                    TB_RecordedRunPath.Text = recordedRunPath ?? string.Empty;
                    if (useRecordedRun && !valid)
                    {
                        TB_RecordedRunPath.BackColor = Color.FromArgb(255, 235, 238);
                        if (!string.IsNullOrEmpty(validationMessage))
                        {
                            toolTip?.SetToolTip(TB_RecordedRunPath, validationMessage);
                        }
                    }
                    else
                    {
                        TB_RecordedRunPath.BackColor = Color.White;
                        toolTip?.SetToolTip(TB_RecordedRunPath, null);
                    }
                }

                if (BT_SelectRecordedRun != null)
                {
                    BT_SelectRecordedRun.Enabled = useRecordedRun;
                }
            }
            finally
            {
                isUpdatingRecordedRunUI = false;
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
                Margin = new Padding(8, 0, 8, 0),
                RowCount = 5,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label header = new Label
            {
                Text = headerText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = ScaleFont("Segoe UI", 9.5F, FontStyle.Bold),
                Padding = new Padding(0, 4, 0, 4)
            };
            layout.Controls.Add(header, 0, 0);

            detailLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = ScaleFont("Segoe UI", 8.5F),
                Padding = new Padding(0, 0, 0, 4),
                Margin = new Padding(0, 0, 0, ScaleSize(4)),
                MaximumSize = new Size(0, ScaleSize(34))
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
            selector.Margin = new Padding(0, ScaleSize(4), 0, ScaleSize(4));
            layout.Controls.Add(selector, 0, 2);

            TableLayoutPanel buttonRow = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, ScaleSize(4), 0, 0),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            buttonRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            buttonRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            connectButton = new Button
            {
                Text = "Connect",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            ConfigureIconButton(connectButton, "\uE71B", $"Connect {context.RoleName.ToLowerInvariant()} camera", minWidth: 44);
            buttonRow.Controls.Add(connectButton, 0, 0);

            captureButton = new Button
            {
                Text = "Capture Preview",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            ConfigureIconButton(captureButton, "\uE114", $"Capture {context.RoleName.ToLowerInvariant()} preview", minWidth: 44);
            buttonRow.Controls.Add(captureButton, 1, 0);

            layout.Controls.Add(buttonRow, 0, 3);

            statusLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(ScaleSize(8), ScaleSize(4), ScaleSize(8), ScaleSize(4)),
                Margin = new Padding(0, ScaleSize(4), 0, 0),
                AutoEllipsis = true,
                MaximumSize = new Size(0, ScaleSize(32))
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
                bool appliedPreferred = false;
                string preferredSerial = GetPreferredSerial(context.Role);
                if (!string.IsNullOrWhiteSpace(preferredSerial))
                {
                    CameraDeviceInfo preferred = cameraDeviceCache.FirstOrDefault(d =>
                        string.Equals(d.SerialNumber, preferredSerial, StringComparison.OrdinalIgnoreCase));
                    if (preferred != null)
                    {
                        context.Selector.SelectedItem = preferred;
                        appliedPreferred = true;
                    }
                }

                if (!appliedPreferred)
                {
                    context.Selector.SelectedIndex = cameraDeviceCache.Count > 0 ? 0 : -1;
                }
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

        private string GetPreferredSerial(CameraRole role)
        {
            if (initSettings == null)
            {
                return null;
            }

            return role == CameraRole.Top ? initSettings.TopCameraSerial : initSettings.FrontCameraSerial;
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
                if (initSettings != null && device != null)
                {
                    if (context.Role == CameraRole.Top)
                    {
                        initSettings.TopCameraSerial = device.SerialNumber;
                    }
                    else
                    {
                        initSettings.FrontCameraSerial = device.SerialNumber;
                    }
                    SaveInitializationSettings();
                }
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
                RunSession session = currentRunSession ?? StartNewRunSession();
                frame = CaptureCameraFrame(topCameraContext, 2000);
                if (frame == null)
                {
                    throw new InvalidOperationException("Top camera returned an empty frame.");
                }

                Directory.CreateDirectory(session.TopFolder);
                string fileName = session.TopImagePath;

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

        private DetectImg LoadRecordedRunTopImage()
        {
            if (!TryGetRecordedTopImagePath(recordedRunPath, out string sourcePath))
            {
                throw new InvalidOperationException("Recorded run top image could not be located.");
            }

            RunSession session = currentRunSession ?? StartNewRunSession();
            Directory.CreateDirectory(session.TopFolder);
            string destination = session.TopImagePath;

            try
            {
                File.Copy(sourcePath, destination, true);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to copy recorded top image: {ex.Message}");
            }

            using (Bitmap bitmap = (Bitmap)Image.FromFile(destination))
            {
                Bitmap previewBitmap = (Bitmap)bitmap.Clone();
                ReplacePictureBoxImage(PB_OriginalImage, previewBitmap);
                Mat mat = BitmapToMat(bitmap);

                DetectImg detectImg = new DetectImg
                {
                    OriImg = mat,
                    OriImgPath = destination,
                    ObjList = new List<ObjectInfo>(),
                    HasResults = false,
                    AttachmentPoints = new List<AttachmentPointInfo>(),
                    AttachmentCenter = new PointF(mat.Width / 2f, mat.Height / 2f),
                    FrontInspections = new List<FrontInspectionResult>(),
                    FrontInspectionComplete = false
                };

                outToLog($"[Recorded] Loaded top image from {sourcePath}.", LogStatus.Info);
                return detectImg;
            }
        }

        private RunSession StartNewRunSession()
        {
            string baseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Runs");
            Directory.CreateDirectory(baseDirectory);

            // Create date folder: YYYY-MM-DD
            DateTime now = DateTime.Now;
            string dateFolder = Path.Combine(baseDirectory, now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(dateFolder);

            // Create Part ID folder with timestamp suffix if needed
            string partID = string.IsNullOrWhiteSpace(currentPartID) ? "Unknown" : SanitizePartID(currentPartID);
            string partFolderName = partID;

            // Check if folder already exists for this Part ID today
            string partFolderPath = Path.Combine(dateFolder, partFolderName);
            if (Directory.Exists(partFolderPath))
            {
                // Add timestamp suffix: PartID_HHmm
                partFolderName = $"{partID}_{now:HHmm}";
                partFolderPath = Path.Combine(dateFolder, partFolderName);
            }

            // Create run folder structure
            Directory.CreateDirectory(partFolderPath);
            Directory.CreateDirectory(Path.Combine(partFolderPath, "Top"));
            Directory.CreateDirectory(Path.Combine(partFolderPath, "Front"));
            Directory.CreateDirectory(Path.Combine(partFolderPath, "Front_Crop"));
            Directory.CreateDirectory(Path.Combine(partFolderPath, "Results"));

            RunSession session = new RunSession(partFolderPath);
            currentRunSession = session;
            outToLog($"[Run] Created run directory: {partFolderPath}", LogStatus.Info);
            return session;
        }

        private string SanitizePartID(string partID)
        {
            if (string.IsNullOrWhiteSpace(partID))
            {
                return "Unknown";
            }

            // Remove invalid file path characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = partID.Trim();
            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }

            // Limit length
            if (sanitized.Length > 50)
            {
                sanitized = sanitized.Substring(0, 50);
            }

            return string.IsNullOrWhiteSpace(sanitized) ? "Unknown" : sanitized;
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

        private static Bitmap CreateCenterCrop500x500(Bitmap source, int? centerX = null, int? centerY = null)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            const int cropWidth = 800;
            const int cropHeight = 900;
            const int defaultCenterX = 2736;
            const int defaultCenterY = 2187;

            int cropCenterX = centerX ?? defaultCenterX;
            int cropCenterY = centerY ?? defaultCenterY;

            int sourceWidth = source.Width;
            int sourceHeight = source.Height;

            // Calculate crop region from center point
            int x = cropCenterX - (cropWidth / 2);
            int y = cropCenterY - (cropHeight / 2);

            // Validate crop fits within image
            if (x < 0 || y < 0 || x + cropWidth > sourceWidth || y + cropHeight > sourceHeight)
            {
                throw new InvalidOperationException(
                    $"Crop region {cropWidth}×{cropHeight} at center ({cropCenterX},{cropCenterY}) doesn't fit in source image {sourceWidth}×{sourceHeight}");
            }

            // Create crop rectangle
            Rectangle cropRect = new Rectangle(x, y, cropWidth, cropHeight);

            // Extract the cropped region
            Bitmap cropped = new Bitmap(cropWidth, cropHeight);
            using (Graphics g = Graphics.FromImage(cropped))
            {
                g.DrawImage(source,
                    new Rectangle(0, 0, cropWidth, cropHeight),  // Destination
                    cropRect,                                     // Source region
                    GraphicsUnit.Pixel);
            }

            return cropped;
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
            UpdateInitSummary();
        }

        private void UpdateInitSummary()
        {
            bool attachmentLoaded = attachmentContext?.IsLoaded == true;
            bool frontAttachmentLoaded = frontAttachmentContext?.IsLoaded == true;
            bool defectLoaded = defectContext?.IsLoaded == true;
            string attachmentLabel = attachmentLoaded ? "Attach OK" : "No Attach";
            string frontAttachmentLabel = frontAttachmentLoaded ? "FrontAttach OK" : "No FrontAttach";
            string defectLabel = defectLoaded ? "Defect OK" : "No Defect";
            bool topConnected = topCameraContext?.IsConnected == true;
            bool frontConnected = frontCameraContext?.IsConnected == true;
            bool turntableConnected = turntableController?.IsConnected == true;
            bool turntableHomed = turntableController?.IsHomed == true;
            double? offset = turntableController?.LastOffsetAngle;
            bool modelsReady = attachmentLoaded && frontAttachmentLoaded && defectLoaded;
            bool camerasReady = topConnected && frontConnected;
            bool turntableReady = turntableConnected && turntableHomed;
            bool recordedReady = true;
            string recordedMessage = string.Empty;
            if (useRecordedRun)
            {
                recordedReady = IsRecordedRunSelectionValid(out recordedMessage);
            }

            string cameraLabel;
            string turntableLabel;
            if (useRecordedRun)
            {
                if (recordedReady)
                {
                    cameraLabel = "Recorded";
                    turntableLabel = "Recorded";
                }
                else
                {
                    cameraLabel = "No Run";
                    turntableLabel = "No Run";
                }
            }
            else
            {
                if (topConnected && frontConnected)
                {
                    cameraLabel = "Cams OK";
                }
                else if (!topConnected && !frontConnected)
                {
                    cameraLabel = "No Cams";
                }
                else if (!topConnected)
                {
                    cameraLabel = "No Top";
                }
                else
                {
                    cameraLabel = "No Front";
                }

                if (!turntableConnected)
                {
                    turntableLabel = "No TT";
                }
                else if (turntableHomed)
                {
                    turntableLabel = offset.HasValue
                        ? $"TT ({offset.Value:0.0}°)"
                        : "TT OK";
                }
                else
                {
                    turntableLabel = "No Home";
                }
            }

            bool ready = modelsReady &&
                (useRecordedRun ? recordedReady : (camerasReady && turntableReady));

            if (LBL_StepModelsStatus != null)
            {
                if (!attachmentLoaded && !frontAttachmentLoaded && !defectLoaded)
                {
                    SetStepStatus(LBL_StepModelsStatus, "Pending", statusNeutralBackground, statusNeutralForeground);
                }
                else if (modelsReady)
                {
                    SetStepStatus(LBL_StepModelsStatus, "Completed", statusPassBackground, statusPassForeground);
                }
                else
                {
                    SetStepStatus(LBL_StepModelsStatus, "Partial", statusFailBackground, statusFailForeground);
                }
            }

            if (LBL_StepCamerasStatus != null)
            {
                if (useRecordedRun)
                {
                    string statusText = recordedReady ? "Recorded run" : "Recorded run (select folder)";
                    SetStepStatus(LBL_StepCamerasStatus, statusText,
                        recordedReady ? statusNeutralBackground : statusFailBackground,
                        recordedReady ? statusNeutralForeground : statusFailForeground);
                }
                else if (!topConnected && !frontConnected)
                {
                    SetStepStatus(LBL_StepCamerasStatus, "Pending", statusNeutralBackground, statusNeutralForeground);
                }
                else if (camerasReady)
                {
                    SetStepStatus(LBL_StepCamerasStatus, "Ready", statusPassBackground, statusPassForeground);
                }
                else
                {
                    SetStepStatus(LBL_StepCamerasStatus, "Assign cameras", statusFailBackground, statusFailForeground);
                }
            }

            if (LBL_StepTurntableStatus != null)
            {
                if (useRecordedRun)
                {
                    string statusText = recordedReady ? "Recorded run" : "Recorded run (select folder)";
                    SetStepStatus(LBL_StepTurntableStatus, statusText,
                        recordedReady ? statusNeutralBackground : statusFailBackground,
                        recordedReady ? statusNeutralForeground : statusFailForeground);
                }
                else if (!turntableConnected)
                {
                    SetStepStatus(LBL_StepTurntableStatus, "Disconnected", statusFailBackground, statusFailForeground);
                }
                else if (turntableHomed)
                {
                    SetStepStatus(LBL_StepTurntableStatus, "Ready", statusPassBackground, statusPassForeground);
                }
                else
                {
                    SetStepStatus(LBL_StepTurntableStatus, "Needs homing", statusFailBackground, statusFailForeground);
                }
            }

            string summaryText = ready
                ? $"Ready: {attachmentLabel} | {defectLabel} | {cameraLabel} | {turntableLabel}"
                : $"{attachmentLabel} | {defectLabel} | {cameraLabel} | {turntableLabel}";
            Color summaryBack = ready ? statusPassBackground : statusNeutralBackground;
            Color summaryFore = ready ? statusPassForeground : statusNeutralForeground;

            if (LBL_InitSummary != null && !LBL_InitSummary.IsDisposed)
            {
                LBL_InitSummary.Text = summaryText;
                LBL_InitSummary.BackColor = summaryBack;
                LBL_InitSummary.ForeColor = summaryFore;
            }

            if (LBL_InitSummaryModal != null && !LBL_InitSummaryModal.IsDisposed)
            {
                LBL_InitSummaryModal.Text = summaryText;
                LBL_InitSummaryModal.BackColor = summaryBack;
                LBL_InitSummaryModal.ForeColor = summaryFore;
            }

            if (BT_InitBeginWorkflow != null && !BT_InitBeginWorkflow.IsDisposed)
            {
                BT_InitBeginWorkflow.Enabled = ready;
            }
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
            else if (!string.IsNullOrWhiteSpace(initSettings?.TurntablePort) &&
                     ports.Contains(initSettings.TurntablePort, StringComparer.OrdinalIgnoreCase))
            {
                string savedPort = ports.First(p =>
                    string.Equals(p, initSettings.TurntablePort, StringComparison.OrdinalIgnoreCase));
                CB_TurntablePort.SelectedItem = savedPort;
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
            UpdateInitSummary();

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
                if (initSettings != null)
                {
                    initSettings.TurntablePort = port;
                    SaveInitializationSettings();
                }

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
                FocusInitializeTab("Connect to the turntable before homing.");
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

                        FocusInitializeTab($"Resolve turntable homing issue: {result.Message}");

                    }

                }

            }

            catch (OperationCanceledException)

            {

                UpdateTurntableStatus("Homing timed out.", statusFailBackground, statusFailForeground);

                outToLog("[Turntable] Homing timed out.", LogStatus.Error);

                FocusInitializeTab("Turntable homing timed out.");

            }

            catch (Exception ex)

            {

                UpdateTurntableStatus($"Homing error: {ex.Message}", statusFailBackground, statusFailForeground);

                outToLog($"[Turntable] Homing error: {ex.Message}", LogStatus.Error);

                FocusInitializeTab($"Turntable homing error: {ex.Message}");

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



        private void FocusInitializeTab(string reason = null)
        {
            ShowInitializationWizard(reason);
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

            // Start cycle timer
            if (cycleTimer == null)
            {
                cycleTimer = new Stopwatch();
            }
            cycleTimer.Restart();
            UpdateCycleTimeDisplay();

            RunSession session = StartNewRunSession();
            if (useRecordedRun && session != null)
            {
                session.RecordedSourcePath = recordedRunPath;
                if (!string.IsNullOrWhiteSpace(recordedRunPath))
                {
                    outToLog($"[Run] Using recorded dataset: {recordedRunPath}", LogStatus.Info);
                }
            }

            DetectImg capturedImage;
            try
            {
                capturedImage = useRecordedRun
                    ? LoadRecordedRunTopImage()
                    : CaptureTopCameraImageForDetection();
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

            // Validate Part ID is entered
            if (string.IsNullOrWhiteSpace(currentPartID))
            {
                outToLog("Please enter a Part ID before running detection.", LogStatus.Warning);
                MessageBox.Show(this, "Please enter a Part/Aligner ID in the text box above the Run Detection button.",
                    "Part ID Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (TB_PartID != null && !TB_PartID.IsDisposed)
                {
                    TB_PartID.Focus();
                }
                return false;
            }

            if (attachmentContext?.IsLoaded != true)
            {
                outToLog("Please load the attachment project before running detection.", LogStatus.Warning);
                FocusInitializeTab("Load the attachment project before running detection.");
                return false;
            }

            if (frontAttachmentContext?.IsLoaded != true)
            {
                outToLog("Please load the front attachment project before running detection.", LogStatus.Warning);
                FocusInitializeTab("Load the front attachment project before running detection.");
                return false;
            }

            if (defectContext?.IsLoaded != true)
            {
                outToLog("Please load the defect project before running detection.", LogStatus.Warning);
                FocusInitializeTab("Load the defect project before running detection.");
                return false;
            }

            if (useRecordedRun)
            {
                if (!IsRecordedRunSelectionValid(out string recordedMessage))
                {
                    string reason = $"Select a saved run before running detection ({recordedMessage}).";
                    outToLog($"Recorded run not ready: {recordedMessage}", LogStatus.Warning);
                    FocusInitializeTab(reason);
                    return false;
                }
            }
            else
            {
                if (topCameraContext?.IsConnected != true)
                {
                    outToLog("Connect the top camera before running detection.", LogStatus.Warning);
                    FocusInitializeTab("Connect the top camera before running detection.");
                    return false;
                }

                if (frontCameraContext?.IsConnected != true)
                {
                    outToLog("Connect the front camera before running detection.", LogStatus.Warning);
                    FocusInitializeTab("Connect the front camera before running detection.");
                    return false;
                }

                if (turntableController?.IsConnected != true)
                {
                    outToLog("Connect to the turntable before running detection.", LogStatus.Warning);
                    FocusInitializeTab("Connect to the turntable before running detection.");
                    return false;
                }

                if (!turntableController.IsHomed)
                {
                    outToLog("Home the turntable before running detection.", LogStatus.Warning);
                    FocusInitializeTab("Home the turntable before running detection.");
                    return false;
                }
            }

            if (string.IsNullOrWhiteSpace(loadedProjectPath))
            {
                outToLog("Attachment project path is missing.", LogStatus.Warning);
                FocusInitializeTab("Attachment project path is missing.");
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

                            // Save top attachments JSON
                            if (currentImage.AttachmentPoints != null && currentImage.AttachmentPoints.Count > 0)
                            {
                                SaveTopAttachmentsJson(currentImage.AttachmentPoints, currentImage.AttachmentCenter, currentImage.OriImgPath);
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
            List<string> snapshot = classNames != null ? new List<string>(classNames) : new List<string>();
            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action(() => HandleClassNameUpdate(snapshot)));
                }
                catch (ObjectDisposedException)
                {
                    // form is closing; ignore
                }
            }
            else
            {
                HandleClassNameUpdate(snapshot);
            }
        }

        private void HandleClassNameUpdate(List<string> classes)
        {
            classNameList = classes ?? new List<string>();
            SyncDefectPolicyWithClasses(classNameList);
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
                    classNames.RemoveAll(name =>
                        name == null ||
                        name.Equals("BackGround", StringComparison.OrdinalIgnoreCase) ||
                        name.Equals("Unknown", StringComparison.OrdinalIgnoreCase) ||
                        name.Equals("Object", StringComparison.OrdinalIgnoreCase));

                    return classNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                }
                catch
                {
                    return new List<string>();
                }
            }

            return new List<string>();
        }

        private void RefreshDefectClassNamesFromModel()
        {
            List<string> classes = GetClassNames() ?? new List<string>();
            HandleClassNameUpdate(classes);
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

            if (useRecordedRun)
            {
                StartRecordedRunProcessing(image);
                return;
            }

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

        private void StartRecordedRunProcessing(DetectImg image)
        {
            if (image == null)
            {
                return;
            }

            if (!IsRecordedRunSelectionValid(out string validationMessage))
            {
                outToLog($"[Sequence] Recorded run unavailable: {validationMessage}", LogStatus.Warning);
                image.FrontInspectionComplete = true;
                UpdateFrontPreviewAndLedger();
                return;
            }

            if (image.AttachmentPoints == null || image.AttachmentPoints.Count == 0)
            {
                image.FrontInspectionComplete = true;
                UpdateFrontPreviewAndLedger();
                return;
            }

            RunSession session = currentRunSession ?? StartNewRunSession();
            List<AttachmentPointInfo> snapshot = image.AttachmentPoints.ToList();
            List<FrontCaptureTicket> tickets = CreateRecordedRunTickets(snapshot);
            if (tickets.Count == 0)
            {
                outToLog("[Sequence] No recorded front images matched the attachment points.", LogStatus.Warning);
                image.FrontInspectionComplete = true;
                UpdateFrontPreviewAndLedger();
                return;
            }

            var cts = new CancellationTokenSource();
            CancellationTokenSource previous = Interlocked.Exchange(ref captureSequenceCts, cts);
            if (previous != null)
            {
                previous.Cancel();
                previous.Dispose();
            }

            outToLog($"[Sequence] Processing recorded run images from {recordedRunPath}.", LogStatus.Progress);

            Task.Run(async () =>
            {
                try
                {
                    List<FrontInspectionResult> inspections = await ProcessFrontCapturesAsync(tickets, cts.Token).ConfigureAwait(false);
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            ApplyFrontInspectionResults(image, inspections);
                            outToLog("[Sequence] Recorded run processing completed.", LogStatus.Success);
                        }));
                    }
                }
                catch (OperationCanceledException)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog("[Sequence] Recorded run processing cancelled.", LogStatus.Warning);
                        }));
                    }
                }
                catch (Exception ex)
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[Sequence] Recorded run processing failed: {ex.Message}", LogStatus.Error);
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

        private List<FrontCaptureTicket> CreateRecordedRunTickets(List<AttachmentPointInfo> points)
        {
            List<FrontCaptureTicket> tickets = new List<FrontCaptureTicket>();
            if (points == null || points.Count == 0)
            {
                return tickets;
            }

            if (string.IsNullOrWhiteSpace(recordedRunPath) || !Directory.Exists(recordedRunPath))
            {
                return tickets;
            }

            RunSession session = currentRunSession ?? StartNewRunSession();
            string sourceFrontFolder = Path.Combine(recordedRunPath, "Front");
            string sourceCropFolder = Path.Combine(recordedRunPath, "Front_Crop");

            if (!Directory.Exists(sourceFrontFolder))
            {
                return tickets;
            }

            // Ensure destination folders exist
            Directory.CreateDirectory(session.FrontFolder);
            Directory.CreateDirectory(Path.Combine(session.Root, "Front_Crop"));

            foreach (AttachmentPointInfo point in points)
            {
                // Try new naming format first: *_idx_NN.png
                string newSearchPattern = $"*_idx_{point.Sequence:00}.png";
                string sourceOriginalPath = Directory.GetFiles(sourceFrontFolder, newSearchPattern)
                    .OrderBy(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();

                // Fall back to old naming format: Front_Index{NN}*.png
                if (string.IsNullOrEmpty(sourceOriginalPath))
                {
                    string oldSearchPattern = $"Front_Index{point.Sequence:00}*.png";
                    sourceOriginalPath = Directory.GetFiles(sourceFrontFolder, oldSearchPattern)
                        .OrderBy(File.GetLastWriteTimeUtc)
                        .FirstOrDefault();
                }

                if (string.IsNullOrEmpty(sourceOriginalPath))
                {
                    outToLog($"[Sequence] Recorded run missing image for index {point.Sequence}.", LogStatus.Warning);
                    continue;
                }

                double angle = TryParseAngleFromFileName(sourceOriginalPath, out double parsed) ? parsed : 0.0;
                string originalFileName = Path.GetFileName(sourceOriginalPath);

                // Copy original to current session's Front folder
                string destOriginalPath = Path.Combine(session.FrontFolder, originalFileName);
                try
                {
                    File.Copy(sourceOriginalPath, destOriginalPath, true);
                }
                catch (Exception ex)
                {
                    outToLog($"[Sequence] Failed to copy recorded image '{sourceOriginalPath}': {ex.Message}", LogStatus.Warning);
                    continue;
                }

                // Note: Dynamic crops will be created during detection based on front attachment position
                outToLog($"[Recorded] Loaded image for index {point.Sequence}, crop will be created during detection.", LogStatus.Info);

                // Create ticket pointing to original image (crop will be created dynamically)
                tickets.Add(new FrontCaptureTicket(point.Sequence, destOriginalPath, destOriginalPath, angle, File.GetLastWriteTime(sourceOriginalPath)));
            }

            return tickets;
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

                    // Save original image to Front folder with new naming: ID_idx_NN.png
                    string filePath = Path.Combine(captureDirectory, $"{currentPartID}_idx_{point.Sequence:00}.png");
                    frame.Save(filePath, ImageFormat.Png);
                    point.CapturedImagePath = filePath;

                    // Note: Crop will be created dynamically during detection based on front attachment position
                    captureTickets.Add(new FrontCaptureTicket(point.Sequence, filePath, filePath, physicalTarget, DateTime.Now));

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
            RunSession session = currentRunSession ?? StartNewRunSession();
            Directory.CreateDirectory(session.FrontFolder);
            return session.FrontFolder;
        }

        private static string BuildAngleToken(double angle)
        {
            string formatted = angle.ToString("+000.0;-000.0;+000.0", CultureInfo.InvariantCulture);
            formatted = formatted.Replace("+", "P").Replace("-", "N").Replace(".", "_");
            return $"Angle_{formatted}";
        }

        private bool TryGetRecordedTopImagePath(string runPath, out string topImagePath)
        {
            topImagePath = null;
            if (string.IsNullOrWhiteSpace(runPath) || !Directory.Exists(runPath))
            {
                return false;
            }

            string candidate = Path.Combine(runPath, "Top.png");
            if (File.Exists(candidate))
            {
                topImagePath = candidate;
                return true;
            }

            string topFolder = Path.Combine(runPath, "Top");
            if (Directory.Exists(topFolder))
            {
                string[] files = Directory.GetFiles(topFolder, "*.png");
                if (files.Length > 0)
                {
                    topImagePath = files.OrderBy(File.GetLastWriteTimeUtc).Last();
                    return true;
                }
            }

            string[] fallback = Directory.GetFiles(runPath, "Top*.png");
            if (fallback.Length > 0)
            {
                topImagePath = fallback.OrderBy(File.GetLastWriteTimeUtc).Last();
                return true;
            }

            return false;
        }

        private static bool TryParseAngleFromFileName(string path, out double angle)
        {
            angle = 0;
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            string fileName = Path.GetFileNameWithoutExtension(path);
            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }

            int idx = fileName.IndexOf("Angle_", StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                return false;
            }

            string token = fileName.Substring(idx + "Angle_".Length);
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            token = token.Replace('P', '+').Replace('N', '-').Replace('_', '.');
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
            {
                angle = parsed;
                return true;
            }

            return false;
        }

        private async Task<List<FrontInspectionResult>> ProcessFrontCapturesAsync(List<FrontCaptureTicket> captures, CancellationToken token)
        {
            if (captures == null || captures.Count == 0 || defectContext?.Process == null || frontAttachmentContext?.Process == null)
            {
                return new List<FrontInspectionResult>();
            }

            List<FrontInspectionResult> results = new List<FrontInspectionResult>(captures.Count);
            List<FrontSequenceData> frontSequences = new List<FrontSequenceData>(); // Collect front attachment data
            await Task.Run(() =>
            {
                foreach (FrontCaptureTicket ticket in captures)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        BeginInvoke(new MethodInvoker(() =>
                        {
                            outToLog($"[FrontAttach] Processing index {ticket.Sequence}...", LogStatus.Progress);
                        }));

                        // Step 1: Load original full-resolution image
                        using (Mat originalMat = CvInvoke.Imread(ticket.OriginalImagePath, ImreadModes.ColorBgr))
                        {
                            if (originalMat == null || originalMat.IsEmpty)
                            {
                                throw new InvalidOperationException($"Unable to load original image: {ticket.OriginalImagePath}");
                            }

                            // Step 2: Detect front attachments in full image
                            List<ObjectInfo> attachments;
                            using (Mat clone = originalMat.Clone())
                            {
                                frontAttachmentContext.Process.Detect(clone, out attachments);
                            }

                            if (attachments == null || attachments.Count == 0)
                            {
                                BeginInvoke(new MethodInvoker(() =>
                                {
                                    outToLog($"[FrontAttach] Index {ticket.Sequence}: No attachments detected - skipping image.", LogStatus.Warning);
                                }));
                                continue;  // Skip this image
                            }

                            // Step 3: Find attachment closest to center X (2736)
                            const int targetCenterX = 2736;
                            ObjectInfo closestAttachment = null;
                            int closestDistance = int.MaxValue;

                            // Collect all attachments for JSON export
                            List<FrontAttachmentData> allAttachments = new List<FrontAttachmentData>();

                            foreach (ObjectInfo attachment in attachments)
                            {
                                int attachCenterX = attachment.DisplayRec.X + attachment.DisplayRec.Width / 2;
                                int attachCenterY = attachment.DisplayRec.Y + attachment.DisplayRec.Height / 2;
                                int distance = Math.Abs(attachCenterX - targetCenterX);

                                if (distance < closestDistance)
                                {
                                    closestDistance = distance;
                                    closestAttachment = attachment;
                                }

                                // Add to collection for JSON export
                                Rectangle bbox = attachment.DisplayRec;
                                float confidence = attachment.confidence != 0 ? attachment.confidence : attachment.classifyScore;
                                string className = string.IsNullOrWhiteSpace(attachment.name) ? "(unknown)" : attachment.name.Trim();

                                allAttachments.Add(new FrontAttachmentData
                                {
                                    Center = new PointData { X = attachCenterX, Y = attachCenterY },
                                    BoundingBox = new BoundingBoxData
                                    {
                                        X = bbox.X,
                                        Y = bbox.Y,
                                        Width = bbox.Width,
                                        Height = bbox.Height
                                    },
                                    ClassName = className,
                                    Confidence = confidence,
                                    DistanceFromTargetCenter = distance,
                                    IsSelectedForCrop = false  // Will update below
                                });
                            }

                            int dynamicCenterX = closestAttachment.DisplayRec.X + closestAttachment.DisplayRec.Width / 2;
                            int dynamicCenterY = closestAttachment.DisplayRec.Y + closestAttachment.DisplayRec.Height / 2;

                            // Mark the selected attachment
                            FrontAttachmentData selectedAttachment = allAttachments.FirstOrDefault(a =>
                                a.DistanceFromTargetCenter == closestDistance);
                            if (selectedAttachment != null)
                            {
                                selectedAttachment.IsSelectedForCrop = true;
                            }

                            BeginInvoke(new MethodInvoker(() =>
                            {
                                outToLog($"[FrontAttach] Index {ticket.Sequence}: Found {attachments.Count} attachment(s). Selected center ({dynamicCenterX}, {dynamicCenterY}), distance from target: {closestDistance}px", LogStatus.Info);
                            }));

                            // Step 4: Create dynamic crop based on detected attachment center
                            Bitmap croppedBitmap;
                            string croppedImagePath;
                            using (Bitmap originalBitmap = originalMat.ToBitmap())
                            {
                                try
                                {
                                    croppedBitmap = CreateCenterCrop500x500(originalBitmap, dynamicCenterX, dynamicCenterY);

                                    // Save dynamic crop to Front_Crop folder
                                    string originalDir = Path.GetDirectoryName(ticket.OriginalImagePath);
                                    string runRoot = Path.GetDirectoryName(originalDir);
                                    string cropDir = Path.Combine(runRoot, "Front_Crop");
                                    Directory.CreateDirectory(cropDir);

                                    string originalFileName = Path.GetFileNameWithoutExtension(ticket.OriginalImagePath);
                                    croppedImagePath = Path.Combine(cropDir, $"{originalFileName}_crop.png");
                                    croppedBitmap.Save(croppedImagePath, ImageFormat.Png);

                                    BeginInvoke(new MethodInvoker(() =>
                                    {
                                        outToLog($"[Crop] Index {ticket.Sequence}: Saved dynamic crop (800×900) centered at ({dynamicCenterX}, {dynamicCenterY}) -> {Path.GetFileName(croppedImagePath)}", LogStatus.Success);
                                    }));
                                }
                                catch (InvalidOperationException ex)
                                {
                                    BeginInvoke(new MethodInvoker(() =>
                                    {
                                        outToLog($"[FrontAttach] Index {ticket.Sequence}: Crop failed at ({dynamicCenterX}, {dynamicCenterY}) - {ex.Message}. Skipping.", LogStatus.Error);
                                    }));
                                    continue;
                                }
                            }

                            // Step 5: Run defect detection on dynamically cropped image
                            List<ObjectInfo> detections;
                            Bitmap overlayBitmap;
                            Bitmap rawBitmap;

                            BeginInvoke(new MethodInvoker(() =>
                            {
                                outToLog($"[Defect] Index {ticket.Sequence}: Running defect detection on dynamic crop...", LogStatus.Progress);
                            }));

                            using (Mat croppedMat = croppedBitmap.ToMat())
                            {
                                // Convert to 3-channel BGR if needed (remove alpha channel)
                                Mat bgrMat = new Mat();
                                if (croppedMat.NumberOfChannels == 4)
                                {
                                    CvInvoke.CvtColor(croppedMat, bgrMat, ColorConversion.Bgra2Bgr);
                                }
                                else if (croppedMat.NumberOfChannels == 1)
                                {
                                    CvInvoke.CvtColor(croppedMat, bgrMat, ColorConversion.Gray2Bgr);
                                }
                                else
                                {
                                    bgrMat = croppedMat.Clone();
                                }

                                // Log detailed Mat information for debugging
                                BeginInvoke(new MethodInvoker(() =>
                                {
                                    outToLog($"[Defect] Index {ticket.Sequence}: Mat info - Original: {croppedMat.NumberOfChannels} channels → Converted: {bgrMat.NumberOfChannels} channels (BGR), Size: {bgrMat.Width}×{bgrMat.Height}", LogStatus.Info);
                                }));

                                using (Mat clone = bgrMat.Clone())
                                {
                                    // Save the exact Mat being sent to defect detection for debugging
                                    string debugImagePath = croppedImagePath.Replace("_crop.png", "_debug_mat.png");
                                    CvInvoke.Imwrite(debugImagePath, clone);

                                    defectContext.Process.Detect(clone, out detections);

                                    BeginInvoke(new MethodInvoker(() =>
                                    {
                                        outToLog($"[Debug] Saved Mat to disk for verification: {Path.GetFileName(debugImagePath)}", LogStatus.Info);
                                    }));
                                }

                                BeginInvoke(new MethodInvoker(() =>
                                {
                                    if (detections != null && detections.Count > 0)
                                    {
                                        outToLog($"[Defect] Index {ticket.Sequence}: Found {detections.Count} defect(s)!", LogStatus.Warning);
                                        foreach (var det in detections)
                                        {
                                            outToLog($"  - {det.name}: confidence={det.confidence:F2}, rect=({det.DisplayRec.X},{det.DisplayRec.Y},{det.DisplayRec.Width},{det.DisplayRec.Height})", LogStatus.Info);
                                        }
                                    }
                                    else
                                    {
                                        outToLog($"[Defect] Index {ticket.Sequence}: No defects detected (model returned 0 results)", LogStatus.Success);
                                    }
                                }));

                                rawBitmap = (Bitmap)croppedBitmap.Clone();
                                using (Mat overlayMat = bgrMat.Clone())
                                {
                                    DrawDefectAnnotations(overlayMat, detections);
                                    overlayBitmap = overlayMat.ToBitmap();
                                }

                                bgrMat.Dispose();
                            }

                            croppedBitmap.Dispose();

                            FrontInspectionResult result = new FrontInspectionResult(ticket.Sequence, croppedImagePath, ticket.AngleDegrees, ticket.CapturedAt, detections, overlayBitmap, rawBitmap);
                            results.Add(result);

                            // Store front sequence data for JSON export
                            frontSequences.Add(new FrontSequenceData
                            {
                                Sequence = ticket.Sequence,
                                ImagePath = Path.GetFileName(ticket.OriginalImagePath),
                                AngleDegrees = ticket.AngleDegrees,
                                AttachmentsDetected = allAttachments,
                                SelectedAttachmentCenter = new PointData { X = dynamicCenterX, Y = dynamicCenterY },
                                CropImagePath = Path.GetFileName(croppedImagePath)
                            });

                            BeginInvoke(new MethodInvoker(() =>
                            {
                                string summary = result.BuildSummary();
                                LogStatus status = result.HasDefects ? LogStatus.Warning : LogStatus.Success;
                                outToLog($"[Defect] Index {result.Sequence}: {summary}", status);
                            }));
                        }
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

            // Save front attachments JSON
            if (frontSequences.Count > 0)
            {
                SaveFrontAttachmentsJson(frontSequences);
            }

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
            if (currentRunSession != null)
            {
                WriteDefectSummaryCsv(sourceImage.FrontInspections);
            }
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

            // Stop cycle timer and log result
            if (cycleTimer != null && cycleTimer.IsRunning)
            {
                cycleTimer.Stop();
                double seconds = cycleTimer.Elapsed.TotalSeconds;
                outToLog($"[Cycle] Total cycle time: {seconds:F2} seconds", LogStatus.Info);
                UpdateCycleTimeDisplay();
            }
        }

        private void UpdateCycleTimeDisplay()
        {
            if (LBL_CycleTime == null || LBL_CycleTime.IsDisposed)
            {
                return;
            }

            if (cycleTimer == null)
            {
                LBL_CycleTime.Text = "Cycle: --";
                return;
            }

            if (cycleTimer.IsRunning)
            {
                LBL_CycleTime.Text = "Cycle: Running...";
                LBL_CycleTime.ForeColor = Color.FromArgb(245, 166, 35); // Orange
            }
            else if (cycleTimer.Elapsed.TotalSeconds > 0)
            {
                double seconds = cycleTimer.Elapsed.TotalSeconds;
                LBL_CycleTime.Text = $"Cycle: {seconds:F2}s";
                LBL_CycleTime.ForeColor = Color.FromArgb(27, 94, 32); // Green
            }
            else
            {
                LBL_CycleTime.Text = "Cycle: --";
                LBL_CycleTime.ForeColor = Color.FromArgb(94, 102, 112); // Gray
            }
        }

        private void WriteDefectSummaryCsv(IEnumerable<FrontInspectionResult> inspections)
        {
            if (currentRunSession == null || inspections == null)
            {
                return;
            }

            try
            {
                Directory.CreateDirectory(currentRunSession.ResultsFolder);
                string csvPath = Path.Combine(currentRunSession.ResultsFolder, "DefectSummary.csv");
                using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    writer.WriteLine("Sequence,AngleDegrees,DefectClass,Confidence,Area,Bounds");
                    foreach (FrontInspectionResult result in inspections.OrderBy(r => r.Sequence))
                    {
                        if (result.Detections != null && result.Detections.Count > 0)
                        {
                            foreach (ObjectInfo detection in result.Detections)
                            {
                                Rectangle rect = detection.DisplayRec;
                                float score = detection.confidence != 0 ? detection.confidence : detection.classifyScore;
                                double area = Math.Max(0, rect.Width) * Math.Max(0, rect.Height);
                                string bounds = $"{rect.Left},{rect.Top},{rect.Width},{rect.Height}";
                                string className = string.IsNullOrWhiteSpace(detection?.name) ? "(unknown)" : detection.name.Trim();
                                writer.WriteLine($"{result.Sequence},{result.AngleDegrees:F2},{className},{score:F4},{area:F0},\"{bounds}\"");
                            }
                        }
                        else
                        {
                            writer.WriteLine($"{result.Sequence},{result.AngleDegrees:F2},PASS,,,");
                        }
                    }
                }
                outToLog($"[Results] Saved defect summary to {Path.Combine(currentRunSession.ResultsFolder, "DefectSummary.csv")}.", LogStatus.Success);
            }
            catch (Exception ex)
            {
                outToLog($"[Results] Failed to write defect summary: {ex.Message}", LogStatus.Warning);
            }
        }

        private void SaveTopAttachmentsJson(List<AttachmentPointInfo> attachmentPoints, PointF imageCenter, string imagePath)
        {
            if (currentRunSession == null || attachmentPoints == null)
            {
                return;
            }

            try
            {
                Directory.CreateDirectory(currentRunSession.TopFolder);

                var jsonData = new TopAttachmentsJson
                {
                    ImagePath = Path.GetFileName(imagePath),
                    ImageCenter = new PointData { X = imageCenter.X, Y = imageCenter.Y },
                    Attachments = new List<TopAttachmentData>(),
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
                };

                foreach (AttachmentPointInfo point in attachmentPoints)
                {
                    Rectangle bbox = point.Source.DisplayRec;
                    float confidence = point.Source.confidence != 0 ? point.Source.confidence : point.Source.classifyScore;
                    string className = string.IsNullOrWhiteSpace(point.Source.name) ? "(unknown)" : point.Source.name.Trim();

                    jsonData.Attachments.Add(new TopAttachmentData
                    {
                        Sequence = point.Sequence,
                        Center = new PointData { X = point.Center.X, Y = point.Center.Y },
                        AngleDegrees = point.AngleDegrees,
                        BoundingBox = new BoundingBoxData
                        {
                            X = bbox.X,
                            Y = bbox.Y,
                            Width = bbox.Width,
                            Height = bbox.Height
                        },
                        ClassName = className,
                        Confidence = confidence
                    });
                }

                string jsonPath = Path.Combine(currentRunSession.TopFolder, "TopAttachments.json");
                string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
                File.WriteAllText(jsonPath, jsonString, Encoding.UTF8);

                outToLog($"[Results] Saved top attachments to {Path.GetFileName(jsonPath)}.", LogStatus.Success);
            }
            catch (Exception ex)
            {
                outToLog($"[Results] Failed to write top attachments JSON: {ex.Message}", LogStatus.Warning);
            }
        }

        private void SaveFrontAttachmentsJson(List<FrontSequenceData> sequences)
        {
            if (currentRunSession == null || sequences == null || sequences.Count == 0)
            {
                return;
            }

            try
            {
                Directory.CreateDirectory(currentRunSession.FrontFolder);

                var jsonData = new FrontAttachmentsJson
                {
                    Sequences = sequences.OrderBy(s => s.Sequence).ToList(),
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
                };

                string jsonPath = Path.Combine(currentRunSession.FrontFolder, "FrontAttachments.json");
                string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
                File.WriteAllText(jsonPath, jsonString, Encoding.UTF8);

                BeginInvoke(new MethodInvoker(() =>
                {
                    outToLog($"[Results] Saved front attachments to {Path.GetFileName(jsonPath)}.", LogStatus.Success);
                }));
            }
            catch (Exception ex)
            {
                BeginInvoke(new MethodInvoker(() =>
                {
                    outToLog($"[Results] Failed to write front attachments JSON: {ex.Message}", LogStatus.Warning);
                }));
            }
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
                Margin = new Padding(8),
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
                    bool hasOverlay = selected.OverlayImage != null;
                    bool hasRaw = selected.RawImage != null;
                    imageToShow = showFrontOverlay ? selected.OverlayImage != null ? (Image)selected.OverlayImage.Clone() : null
                                                   : selected.RawImage != null ? (Image)selected.RawImage.Clone() : null;

                    // Debug logging
                    if (selected.Detections != null && selected.Detections.Count > 0)
                    {
                        outToLog($"[Preview] Displaying index {selected.Sequence}: ShowOverlay={showFrontOverlay}, HasOverlay={hasOverlay}, HasRaw={hasRaw}, DetectionCount={selected.Detections.Count}", LogStatus.Info);
                    }
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
                string taskLabel = string.IsNullOrWhiteSpace(detail) ? "process" : detail.Trim();
                description = $"Task '{taskLabel}' started.";
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
            Font baseFont = ScaleFont("Segoe UI", 9F);
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
            if (groupWorkflowInit != null)
            {
                StyleGroupSurface(groupWorkflowInit, surface, border);
            }
            if (groupInitAttachment != null)
            {
                StyleGroupSurface(groupInitAttachment, surface, border);
            }
            if (groupInitDefect != null)
            {
                StyleGroupSurface(groupInitDefect, surface, border);
            }

            if (leftTabs != null)
            {
                leftTabs.Font = ScaleFont("Segoe UI Semibold", 9F);
                leftTabs.Appearance = TabAppearance.Normal;
                leftTabs.DrawMode = TabDrawMode.Normal;
                leftTabs.ItemSize = new Size(ScaleSize(120), ScaleSize(28));
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

            foreach (Button initButton in new[] { BT_InitAttachmentBrowse, BT_InitAttachmentLoad, BT_InitDefectBrowse, BT_InitDefectLoad, BT_CameraRefresh, BT_TopCameraConnect, BT_TopCameraCapture, BT_FrontCameraConnect, BT_FrontCameraCapture, BT_TurntableRefresh, BT_TurntableConnect, BT_TurntableHome, BT_InitBeginWorkflow, BT_OpenInitWizard })
            {
                if (initButton != null)
                {
                    StylePrimaryButton(initButton, accentPrimary, accentHover);
                    if (initButton == BT_TurntableHome)
                    {
                        initButton.AutoSize = true;
                        initButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                        initButton.MinimumSize = new Size(0, 0);
                        initButton.Margin = new Padding(ScaleSize(12), ScaleSize(4), 0, ScaleSize(4));
                    }
                    else
                    {
                        initButton.AutoSize = false;
                        initButton.Height = ScaleSize(28);
                        initButton.Margin = new Padding(ScaleSize(4));
                        initButton.MinimumSize = new Size(ScaleSize(120), ScaleSize(26));
                    }
                }
            }

            if (BT_CameraRefresh != null)
            {
                BT_CameraRefresh.MinimumSize = new Size(ScaleSize(110), ScaleSize(30));
                BT_CameraRefresh.Height = ScaleSize(32);
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
            RTB_Info.Font = ScaleFont("Consolas", 9F);
            RTB_Info.Margin = new Padding(ScaleSize(16));
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
            TV_Logic.Font = ScaleFont("Segoe UI", 9F);
            TV_Logic.LineColor = Color.FromArgb(210, 216, 226);

            foreach (ComboBox combo in new[] { CB_GroupOperator, CB_Field, CB_Operator, CB_Value, CB_TurntablePort })
            {
                combo.FlatStyle = FlatStyle.Flat;
                combo.BackColor = Color.FromArgb(252, 253, 255);
                combo.ForeColor = accentSecondary;
                combo.Margin = new Padding(0, ScaleSize(4), 0, ScaleSize(4));
            }

            CB_Value.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            CB_Value.AutoCompleteSource = AutoCompleteSource.ListItems;
            CB_Value.DropDownStyle = ComboBoxStyle.DropDown;

            StylePrimaryButton(BT_AddRule, accentPrimary, accentHover);
            StylePrimaryButton(BT_AddGroup, accentPrimary, accentHover);
            BT_AddRule.Height = ScaleSize(28);
            BT_AddGroup.Height = ScaleSize(28);
            BT_AddRule.Margin = new Padding(ScaleSize(2), 0, ScaleSize(2), 0);
            BT_AddGroup.Margin = new Padding(ScaleSize(2), 0, ScaleSize(2), 0);

            StylePrimaryButton(BT_LogicHelp, accentSecondary, Color.FromArgb(76, 84, 94));
            BT_LogicHelp.BackColor = Color.FromArgb(231, 233, 238);
            BT_LogicHelp.ForeColor = Color.FromArgb(58, 66, 78);
            BT_LogicHelp.Width = ScaleSize(28);
            BT_LogicHelp.Height = ScaleSize(28);

            StyleDangerButton(BT_RemoveNode, Color.FromArgb(224, 68, 67), Color.FromArgb(205, 58, 57));
            BT_RemoveNode.Height = ScaleSize(28);
            BT_RemoveNode.Margin = new Padding(ScaleSize(2), 0, ScaleSize(2), 0);

            StyleStatusLabel(LBL_ResultStatus);
            StyleStatusLabel(LBL_WorkflowStatus);
            StyleStatusLabel(LBL_InitSummary);
            StyleStatusLabel(LBL_TurntableStatus);
            if (LBL_ResultStatus != null)
            {
                LBL_ResultStatus.Margin = new Padding(0);
                LBL_ResultStatus.AutoSize = true;
                LBL_ResultStatus.Dock = DockStyle.Top;
            }
            if (LBL_WorkflowStatus != null)
            {
                LBL_WorkflowStatus.Margin = new Padding(0);
                LBL_WorkflowStatus.AutoSize = true;
                LBL_WorkflowStatus.Dock = DockStyle.Top;
            }
            SetLogicStatusNeutral("Awaiting defect inspection.");

            this.Padding = new Padding(ScaleSize(12));
        }

        private void BuildLogicBuilderUI()
        {
            toolTip = new ToolTip();

            EnsureInitLayout();

            if (leftTabs == null || leftTabs.IsDisposed)
            {
                leftTabs = new TabControl
                {
                    Name = "leftTabs"
                };
            }

            leftTabs.Dock = DockStyle.Fill;
            leftTabs.Appearance = TabAppearance.Normal;
            leftTabs.SizeMode = TabSizeMode.Fixed;
            leftTabs.ItemSize = new Size(ScaleSize(120), ScaleSize(28));
            leftTabs.Margin = new Padding(0, 0, ScaleSize(6), 0);
            leftTabs.Multiline = true;

            if (!tableLayoutPanelMain.Controls.Contains(leftTabs))
            {
                tableLayoutPanelMain.Controls.Add(leftTabs, 0, 0);
                tableLayoutPanelMain.SetColumn(leftTabs, 0);
                tableLayoutPanelMain.SetRow(leftTabs, 0);
            }

            if (tabWorkflow == null || tabWorkflow.IsDisposed)
            {
                tabWorkflow = new TabPage("Workflow");
            }
            tabWorkflow.Text = "Workflow";
            tabWorkflow.Padding = new Padding(ScaleSize(8));
            tabWorkflow.Controls.Clear();

            if (panelStepsHost != null && panelStepsHost.Parent != null && panelStepsHost.Parent != tabWorkflow)
            {
                panelStepsHost.Parent.Controls.Remove(panelStepsHost);
            }
            if (panelStepsHost != null)
            {
                panelStepsHost.Dock = DockStyle.Fill;
            }

            if (tabLogicBuilder == null || tabLogicBuilder.IsDisposed)
            {
                tabLogicBuilder = new TabPage("Logic Builder");
            }
            tabLogicBuilder.Text = "Logic Builder";
            tabLogicBuilder.Padding = new Padding(ScaleSize(12));
            tabLogicBuilder.Controls.Clear();

            workflowLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            workflowLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            workflowLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            workflowLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            LBL_WorkflowStatus = new Label
            {
                AutoSize = true,
                Text = "Awaiting project.",
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Top
            };

            Panel workflowFooter = new Panel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(ScaleSize(16), ScaleSize(8), ScaleSize(16), ScaleSize(8))
            };
            workflowFooter.Controls.Add(LBL_WorkflowStatus);

            Control workflowBody = panelStepsHost ?? tableLayoutSteps;
            if (workflowBody != null && workflowBody.Parent != null && workflowBody.Parent != workflowLayout)
            {
                workflowBody.Parent.Controls.Remove(workflowBody);
            }
            if (workflowBody == tableLayoutSteps)
            {
                tableLayoutSteps.Dock = DockStyle.Top;
            }
            workflowLayout.Controls.Add(workflowBody, 0, 0);
            workflowLayout.Controls.Add(workflowFooter, 0, 1);
            tabWorkflow.Controls.Add(workflowLayout);

            if (panelStepsHost != null)
            {
                Size preferredSize = tableLayoutSteps?.PreferredSize ?? Size.Empty;
                panelStepsHost.AutoScrollMinSize = new Size(0, preferredSize.Height + ScaleSize(40));
            }

            groupClassRules = new GroupBox
            {
                Text = "Defect Class Policy",
                Dock = DockStyle.Fill,
                Padding = new Padding(ScaleSize(12), ScaleSize(24), ScaleSize(12), ScaleSize(12)),
                MinimumSize = new Size(ScaleSize(420), ScaleSize(0)),
                AutoSize = false
            };

            flowClassRules = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Margin = new Padding(ScaleSize(4)),
                Padding = new Padding(ScaleSize(4))
            };
            groupClassRules.Controls.Add(flowClassRules);

            groupLogic = new GroupBox
            {
                Text = "Logic Rules",
                Dock = DockStyle.Fill,
                Padding = new Padding(ScaleSize(12), ScaleSize(24), ScaleSize(12), ScaleSize(12))
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
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleSize(32)));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleSize(32)));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleSize(32)));
            tableLayoutLogicEditor.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleSize(32)));
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
                Margin = new Padding(4, 0, 4, 0)
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
                RowCount = 3
            };
            logicRootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56F));
            logicRootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 48F));
            logicRootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 52F));
            logicRootLayout.Controls.Add(logicHeaderPanel, 0, 0);
            logicRootLayout.Controls.Add(groupClassRules, 0, 1);
            logicRootLayout.Controls.Add(groupLogic, 0, 2);

            tabLogicBuilder.Controls.Add(logicRootLayout);
            RefreshClassRuleCardsUI();

            leftTabs.TabPages.Clear();
            leftTabs.TabPages.Add(tabWorkflow);
            leftTabs.TabPages.Add(tabLogicBuilder);
            leftTabs.SelectedTab = tabWorkflow;
        }

        private TableLayoutPanel BuildInitializationLayout()
        {
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(246, 248, 252),
                Padding = new Padding(8, 8, 8, 12),
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            int row = 0;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(BuildInitializationHeader(), 0, row++);

            Control modelBody = BuildModelStepBody();
            Panel modelsCard = CreateStepCard("Step 1 - Load Models", "Select the Attachment and Defect TSP projects used for this workstation.", modelBody, out LBL_StepModelsStatus);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(modelsCard, 0, row++);

            Control cameraBody = BuildCameraStepBody();
            Panel camerasCard = CreateStepCard("Step 2 - Configure Cameras", "Assign top and front cameras, then capture a preview to confirm connectivity.", cameraBody, out LBL_StepCamerasStatus);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(camerasCard, 0, row++);

            Control turntableBody = BuildTurntableStepBody();
            Panel turntableCard = CreateStepCard("Step 3 - Prepare Turntable", "Connect to the turntable controller and perform homing to establish zero.", turntableBody, out LBL_StepTurntableStatus);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(turntableCard, 0, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(BuildSummarySection(), 0, row++);

            return layout;
        }

        private void BuildWorkflowInitializationCard()
        {
            if (tableLayoutSteps == null)
            {
                return;
            }

            groupWorkflowInit = new GroupBox
            {
                Text = "Initialize System",
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 12, 12, 12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(0, ScaleSize(140))
            };

            TableLayoutPanel inner = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            inner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inner.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            inner.RowCount = 3;
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            LBL_InitSummary = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(0, 4, 12, 4),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(94, 102, 112),
                Height = ScaleSize(32),
                AutoEllipsis = true,
                MaximumSize = new Size(0, ScaleSize(32))
            };

            if (BT_OpenInitWizard == null || BT_OpenInitWizard.IsDisposed)
            {
                BT_OpenInitWizard = new Button();
            }
            BT_OpenInitWizard.Text = "Open Wizard";
            BT_OpenInitWizard.AutoSize = true;
            BT_OpenInitWizard.Margin = new Padding(0);
            BT_OpenInitWizard.Anchor = AnchorStyles.Right;
            BT_OpenInitWizard.MinimumSize = new Size(ScaleSize(110), ScaleSize(32));
            BT_OpenInitWizard.Height = ScaleSize(32);
            BT_OpenInitWizard.Click -= BT_OpenInitWizard_Click;
            BT_OpenInitWizard.Click += BT_OpenInitWizard_Click;

            inner.Controls.Add(LBL_InitSummary, 0, 0);
            inner.Controls.Add(BT_OpenInitWizard, 1, 0);

            if (CHK_UseRecordedRun == null || CHK_UseRecordedRun.IsDisposed)
            {
                CHK_UseRecordedRun = new CheckBox();
            }
            CHK_UseRecordedRun.AutoSize = true;
            CHK_UseRecordedRun.Text = "Use saved run images";
            CHK_UseRecordedRun.Margin = new Padding(0, ScaleSize(4), ScaleSize(8), 0);
            CHK_UseRecordedRun.CheckedChanged -= CHK_UseRecordedRun_CheckedChanged;
            CHK_UseRecordedRun.CheckedChanged += CHK_UseRecordedRun_CheckedChanged;

            if (TB_RecordedRunPath == null || TB_RecordedRunPath.IsDisposed)
            {
                TB_RecordedRunPath = new TextBox();
            }
            TB_RecordedRunPath.ReadOnly = true;
            TB_RecordedRunPath.Margin = new Padding(0, ScaleSize(2), ScaleSize(8), 0);
            TB_RecordedRunPath.MinimumSize = new Size(ScaleSize(220), ScaleSize(28));
            TB_RecordedRunPath.Height = ScaleSize(28);
            TB_RecordedRunPath.Dock = DockStyle.Fill;

            if (BT_SelectRecordedRun == null || BT_SelectRecordedRun.IsDisposed)
            {
                BT_SelectRecordedRun = new Button();
            }
            BT_SelectRecordedRun.AutoSize = true;
            BT_SelectRecordedRun.Text = "Browse...";
            BT_SelectRecordedRun.Margin = new Padding(0, 0, 0, 0);
            BT_SelectRecordedRun.MinimumSize = new Size(ScaleSize(90), ScaleSize(28));
            BT_SelectRecordedRun.Height = ScaleSize(28);
            BT_SelectRecordedRun.Click -= BT_SelectRecordedRun_Click;
            BT_SelectRecordedRun.Click += BT_SelectRecordedRun_Click;

            recordedRunPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, ScaleSize(6), 0, 0),
                Padding = new Padding(0),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            CHK_UseRecordedRun.Margin = new Padding(0, ScaleSize(2), ScaleSize(12), 0);
            TB_RecordedRunPath.Margin = new Padding(0, 0, ScaleSize(8), 0);
            BT_SelectRecordedRun.Margin = new Padding(0);

            recordedRunPanel.Controls.Add(CHK_UseRecordedRun);
            recordedRunPanel.Controls.Add(TB_RecordedRunPath);
            recordedRunPanel.Controls.Add(BT_SelectRecordedRun);

            inner.Controls.Add(recordedRunPanel, 0, 1);
            inner.SetColumnSpan(recordedRunPanel, 2);

            recordedRunPanel.SizeChanged += (s, e) => AdjustRecordedRunLayout();
            inner.SizeChanged += (s, e) => AdjustRecordedRunLayout();

            groupWorkflowInit.Controls.Add(inner);

            tableLayoutSteps.SuspendLayout();

            // Rebuild layout order with initialization group inserted before step 4
            tableLayoutSteps.Controls.Clear();
            tableLayoutSteps.RowStyles.Clear();
            tableLayoutSteps.RowCount = 7;
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutSteps.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            tableLayoutSteps.Controls.Add(groupStep1, 0, 0);
            tableLayoutSteps.Controls.Add(groupStep2, 0, 1);
            tableLayoutSteps.Controls.Add(groupWorkflowInit, 0, 2);
            tableLayoutSteps.Controls.Add(groupStep4, 0, 3);
            tableLayoutSteps.Controls.Add(groupGallery, 0, 4);
            tableLayoutSteps.Controls.Add(groupStep5, 0, 5);
            tableLayoutSteps.Controls.Add(groupStep7, 0, 6);

            tableLayoutSteps.ResumeLayout();

            ApplyResponsiveLayout(force: true);
            UpdateInitSummary();
            AdjustRecordedRunLayout();
        }

        private Control BuildInitializationHeader()
        {
            FlowLayoutPanel headerFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(ScaleSize(8), 0, ScaleSize(8), ScaleSize(12)),
                Margin = new Padding(0, 0, 0, ScaleSize(12))
            };

            Label title = new Label
            {
                AutoSize = true,
                Text = "System Initialization",
                Font = ScaleFont("Segoe UI Semibold", 16F),
                ForeColor = Color.FromArgb(47, 57, 72),
                Margin = new Padding(0, 0, 0, ScaleSize(4))
            };

            Label subtitle = new Label
            {
                AutoSize = true,
                Text = "Complete the following steps to prepare models, cameras, and the turntable before running inspections.",
                Font = ScaleFont("Segoe UI", 9.5F),
                ForeColor = Color.FromArgb(94, 102, 112),
                Margin = new Padding(0, 0, 0, 0),
                MaximumSize = new Size(900, 0)
            };

            LBL_InitPrompt = new Label
            {
                AutoSize = true,
                Visible = false,
                ForeColor = Color.FromArgb(178, 34, 34),
                Font = ScaleFont("Segoe UI", 9F, FontStyle.Italic),
                Margin = new Padding(0, ScaleSize(8), 0, 0),
                MaximumSize = new Size(900, 0)
            };

            headerFlow.Controls.Add(title);
            headerFlow.Controls.Add(subtitle);
            headerFlow.Controls.Add(LBL_InitPrompt);

            return headerFlow;
        }

        private Panel CreateStepCard(string title, string description, Control body, out Label statusChip)
        {
            statusChip = CreateStatusChip("Pending", statusNeutralBackground, statusNeutralForeground);
            StyleStepStatusChip(statusChip);

            Label titleLabel = new Label
            {
                AutoSize = true,
                Text = title,
                Font = ScaleFont("Segoe UI Semibold", 11F),
                ForeColor = Color.FromArgb(47, 57, 72),
                Margin = new Padding(0, 0, 0, 2)
            };

            Label descriptionLabel = new Label
            {
                AutoSize = true,
                Text = description,
                Font = ScaleFont("Segoe UI", 9.5F),
                ForeColor = Color.FromArgb(94, 102, 112),
                Margin = new Padding(0, 0, 0, 8),
                MaximumSize = new Size(900, 0)
            };

            TableLayoutPanel header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Margin = new Padding(0, 0, 0, ScaleSize(8))
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.Controls.Add(titleLabel, 0, 0);
            header.Controls.Add(statusChip, 1, 0);
            header.Controls.Add(descriptionLabel, 0, 1);
            header.SetColumnSpan(descriptionLabel, 2);

            Panel card = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(ScaleSize(16)),
                Margin = new Padding(0, 0, 0, ScaleSize(12))
            };

            if (body != null)
            {
                body.Dock = DockStyle.Top;
                card.Controls.Add(body);
            }
            card.Controls.Add(header);
            header.Dock = DockStyle.Top;

            return card;
        }

        private Label CreateStatusChip(string text, Color background, Color foreground)
        {
            return new Label
            {
                AutoSize = true,
                Text = text,
                BackColor = background,
                ForeColor = foreground,
                Font = ScaleFont("Segoe UI Semibold", 8.5F),
                Padding = new Padding(ScaleSize(10), ScaleSize(4), ScaleSize(10), ScaleSize(4)),
                Margin = new Padding(ScaleSize(12), 0, 0, 0)
            };
        }

        private Control BuildModelStepBody()
        {
            TableLayoutPanel grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            groupInitAttachment = CreateModelGroup(
                "Attachment Model (Top View)",
                BT_InitAttachmentBrowse_Click,
                BT_InitAttachmentLoad_Click,
                out TB_InitAttachmentPath,
                out BT_InitAttachmentBrowse,
                out BT_InitAttachmentLoad,
                out LBL_InitAttachmentStatus);
            groupInitAttachment.Margin = new Padding(0, 8, 0, 4);

            groupInitFrontAttachment = CreateModelGroup(
                "Front Attachment Model",
                BT_InitFrontAttachmentBrowse_Click,
                BT_InitFrontAttachmentLoad_Click,
                out TB_InitFrontAttachmentPath,
                out BT_InitFrontAttachmentBrowse,
                out BT_InitFrontAttachmentLoad,
                out LBL_InitFrontAttachmentStatus);
            groupInitFrontAttachment.Margin = new Padding(0, 4, 0, 4);

            groupInitDefect = CreateModelGroup(
                "Defect Model",
                BT_InitDefectBrowse_Click,
                BT_InitDefectLoad_Click,
                out TB_InitDefectPath,
                out BT_InitDefectBrowse,
                out BT_InitDefectLoad,
                out LBL_InitDefectStatus);
            groupInitDefect.Margin = new Padding(0, 4, 0, 0);

            grid.Controls.Add(groupInitAttachment, 0, 0);
            grid.Controls.Add(groupInitFrontAttachment, 0, 1);
            grid.Controls.Add(groupInitDefect, 0, 2);

            return grid;
        }

        private GroupBox CreateModelGroup(
            string title,
            EventHandler browseHandler,
            EventHandler loadHandler,
            out TextBox pathDisplay,
            out Button browseButton,
            out Button loadButton,
            out Label statusLabel)
        {
            GroupBox group = new GroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12, 16, 12, 12)
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                ColumnCount = 3,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            pathDisplay = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Margin = new Padding(0, 0, 0, 4)
            };
            layout.Controls.Add(pathDisplay, 0, 0);
            layout.SetColumnSpan(pathDisplay, 3);

            browseButton = new Button
            {
                Text = "Browse...",
                Dock = DockStyle.Fill
            };
            browseButton.Click += browseHandler;
            ConfigureIconButton(browseButton, "\uE8B7", $"Browse {title} project");
            layout.Controls.Add(browseButton, 1, 1);

            loadButton = new Button
            {
                Text = "Load",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            loadButton.Click += loadHandler;
            ConfigureIconButton(loadButton, "\uE74D", $"Load {title} project");
            layout.Controls.Add(loadButton, 2, 1);

            statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Not loaded.",
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 4, 8, 4)
            };
            layout.Controls.Add(statusLabel, 0, 2);
            layout.SetColumnSpan(statusLabel, 3);

            group.Controls.Add(layout);
            return group;
        }

        private Control BuildCameraStepBody()
        {
            TableLayoutPanel container = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Margin = new Padding(0)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            BT_CameraRefresh = new Button
            {
                Text = "Refresh Devices",
                AutoSize = true,
                Margin = new Padding(0, 0, 12, 0)
            };
            BT_CameraRefresh.Click += BT_CameraRefresh_Click;
            ConfigureIconButton(BT_CameraRefresh, "\uE72C", "Refresh camera list", minWidth: 40, fontSize: 16f, customMargin: new Padding(0, 0, ScaleSize(12), 0));
            toolbar.Controls.Add(BT_CameraRefresh);

            Label hint = new Label
            {
                AutoSize = true,
                Text = "Assign top and front cameras, then capture a preview to verify the feed.",
                ForeColor = Color.FromArgb(94, 102, 112),
                Margin = new Padding(0, 6, 0, 0)
            };
            toolbar.Controls.Add(hint);

            container.Controls.Add(toolbar, 0, 0);

            TableLayoutPanel cameraColumns = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 12, 0, 0)
            };
            cameraColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            cameraColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            Label topDetail;
            TableLayoutPanel topLayout = CreateCameraRoleLayout(
                topCameraContext,
                "Top Camera",
                out CB_TopCameraSelect,
                out BT_TopCameraConnect,
                out BT_TopCameraCapture,
                out topDetail,
                out LBL_TopCameraStatus);

            Label frontDetail;
            TableLayoutPanel frontLayout = CreateCameraRoleLayout(
                frontCameraContext,
                "Front Camera",
                out CB_FrontCameraSelect,
                out BT_FrontCameraConnect,
                out BT_FrontCameraCapture,
                out frontDetail,
                out LBL_FrontCameraStatus);

            cameraColumns.Controls.Add(topLayout, 0, 0);
            cameraColumns.Controls.Add(frontLayout, 1, 0);

            container.Controls.Add(cameraColumns, 0, 1);

            topCameraContext.Selector = CB_TopCameraSelect;
            topCameraContext.ConnectButton = BT_TopCameraConnect;
            topCameraContext.CaptureButton = BT_TopCameraCapture;
            topCameraContext.StatusLabel = LBL_TopCameraStatus;
            topCameraContext.DetailLabel = topDetail;

            frontCameraContext.Selector = CB_FrontCameraSelect;
            frontCameraContext.ConnectButton = BT_FrontCameraConnect;
            frontCameraContext.CaptureButton = BT_FrontCameraCapture;
            frontCameraContext.StatusLabel = LBL_FrontCameraStatus;
            frontCameraContext.DetailLabel = frontDetail;

            CB_TopCameraSelect.SelectedIndexChanged += (s, e) => UpdateCameraSelectionChanged(topCameraContext);
            CB_FrontCameraSelect.SelectedIndexChanged += (s, e) => UpdateCameraSelectionChanged(frontCameraContext);
            BT_TopCameraConnect.Click += (s, e) => ToggleCameraConnection(topCameraContext);
            BT_FrontCameraConnect.Click += (s, e) => ToggleCameraConnection(frontCameraContext);
            BT_TopCameraCapture.Click += (s, e) => CaptureCameraPreview(topCameraContext);
            BT_FrontCameraCapture.Click += (s, e) => CaptureCameraPreview(frontCameraContext);

            return container;
        }

        private Control BuildTurntableStepBody()
        {
            TableLayoutPanel layout = new TableLayoutPanel
            {
                ColumnCount = 4,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layout.RowCount = 2;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            CB_TurntablePort = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            layout.Controls.Add(CB_TurntablePort, 0, 0);

            BT_TurntableRefresh = new Button
            {
                Text = "Refresh",
                Dock = DockStyle.Fill,
                Margin = new Padding(8, 0, 0, 0)
            };
            BT_TurntableRefresh.Click += BT_TurntableRefresh_Click;
            ConfigureIconButton(BT_TurntableRefresh, "\uE72C", "Refresh COM ports", customMargin: new Padding(8, 0, 0, 0));
            layout.Controls.Add(BT_TurntableRefresh, 1, 0);

            BT_TurntableConnect = new Button
            {
                Text = "Connect",
                Dock = DockStyle.Fill,
                Margin = new Padding(8, 0, 0, 0)
            };
            BT_TurntableConnect.Click += BT_TurntableConnect_Click;
            ConfigureIconButton(BT_TurntableConnect, "\uE71B", "Connect turntable", customMargin: new Padding(8, 0, 0, 0));
            layout.Controls.Add(BT_TurntableConnect, 2, 0);

            BT_TurntableHome = new Button
            {
                Text = "Home",
                Dock = DockStyle.Fill,
                Enabled = false,
                Margin = new Padding(8, 0, 0, 0)
            };
            BT_TurntableHome.Click += BT_TurntableHome_Click;
            ConfigureIconButton(BT_TurntableHome, "\uE80F", "Return to home position", customMargin: new Padding(8, 0, 0, 0));
            layout.Controls.Add(BT_TurntableHome, 3, 0);

            LBL_TurntableStatus = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Disconnected.",
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 4, 8, 4),
                Margin = new Padding(0, 8, 0, 0)
            };
            layout.Controls.Add(LBL_TurntableStatus, 0, 1);
            layout.SetColumnSpan(LBL_TurntableStatus, 4);

            return layout;
        }

        private Control BuildSummarySection()
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(16, 16, 16, 16),
                Margin = new Padding(0)
            };

            FlowLayoutPanel flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0)
            };

            LBL_InitSummaryModal = new Label
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 8, 8, 8),
                Margin = new Padding(0, 0, 16, 0)
            };

            BT_InitBeginWorkflow = new Button
            {
                Text = "Begin Workflow",
                AutoSize = true,
                Enabled = false,
                Margin = new Padding(0)
            };
            BT_InitBeginWorkflow.Click += BT_InitBeginWorkflow_Click;

            flow.Controls.Add(LBL_InitSummaryModal);
            flow.Controls.Add(BT_InitBeginWorkflow);

            panel.Controls.Add(flow);

            return panel;
        }

        private void BT_OpenInitWizard_Click(object sender, EventArgs e)
        {
            ShowInitializationWizard();
        }

        private void CHK_UseRecordedRun_CheckedChanged(object sender, EventArgs e)
        {
            if (isUpdatingRecordedRunUI)
            {
                return;
            }

            useRecordedRun = CHK_UseRecordedRun?.Checked == true;
            if (initSettings != null)
            {
                initSettings.UseRecordedRun = useRecordedRun;
                SaveInitializationSettings();
            }

            UpdateRecordedRunUiState();
            UpdateInitSummary();
        }

        private void BT_SelectRecordedRun_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a saved run folder";
                if (!string.IsNullOrWhiteSpace(recordedRunPath) && Directory.Exists(recordedRunPath))
                {
                    dialog.SelectedPath = recordedRunPath;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    recordedRunPath = dialog.SelectedPath;
                    if (initSettings != null)
                    {
                        initSettings.RecordedRunPath = recordedRunPath;
                        SaveInitializationSettings();
                    }
                    UpdateRecordedRunUiState();
                    UpdateInitSummary();
                }
            }
        }

        private void ShowInitializationWizard(string reason = null)
        {
            EnsureInitLayout();

            if (LBL_InitPrompt != null && !LBL_InitPrompt.IsDisposed)
            {
                if (string.IsNullOrWhiteSpace(reason))
                {
                    LBL_InitPrompt.Visible = false;
                }
                else
                {
                    LBL_InitPrompt.Text = reason;
                    LBL_InitPrompt.Visible = true;
                }
            }

            UpdateInitSummary();

            if (activeInitDialog != null && !activeInitDialog.IsDisposed)
            {
                activeInitDialog.BringToFront();
                activeInitDialog.Activate();
                return;
            }

            if (BT_InitBeginWorkflow != null && !BT_InitBeginWorkflow.IsDisposed)
            {
                BT_InitBeginWorkflow.DialogResult = DialogResult.OK;
            }

            using (InitializationDialog dialog = new InitializationDialog(initLayout))
            {
                activeInitDialog = dialog;
                if (BT_InitBeginWorkflow != null && !BT_InitBeginWorkflow.IsDisposed)
                {
                    dialog.AcceptButton = BT_InitBeginWorkflow;
                }

                dialog.ShowDialog(this);
                activeInitDialog = null;
            }

            UpdateInitSummary();
        }

        private void BT_InitBeginWorkflow_Click(object sender, EventArgs e)
        {
            if (!IsInitializationComplete())
            {
                MessageBox.Show(
                    this,
                    "Please complete all initialization steps before beginning the workflow.",
                    "Initialization Incomplete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                UpdateInitSummary();
                return;
            }

            if (activeInitDialog != null)
            {
                activeInitDialog.DialogResult = DialogResult.OK;
                activeInitDialog.Close();
            }
        }

        private bool IsInitializationComplete()
        {
            return attachmentContext?.IsLoaded == true
                && defectContext?.IsLoaded == true
                && topCameraContext?.IsConnected == true
                && frontCameraContext?.IsConnected == true
                && turntableController?.IsConnected == true
                && turntableController?.IsHomed == true;
        }

        private void SetStepStatus(Label label, string text, Color background, Color foreground)
        {
            if (label == null || label.IsDisposed)
            {
                return;
            }

            StyleStepStatusChip(label);
            label.Text = text;
            label.BackColor = background;
            label.ForeColor = foreground;
        }

        private void StyleStepStatusChip(Label label)
        {
            if (label == null || label.IsDisposed)
            {
                return;
            }

            label.Font = ScaleFont("Segoe UI Semibold", 8.5F);
            label.Padding = new Padding(ScaleSize(10), ScaleSize(4), ScaleSize(10), ScaleSize(4));
            label.Margin = new Padding(ScaleSize(12), 0, 0, 0);
        }

        private sealed class InitializationSettings
        {
            public string AttachmentPath { get; set; }
            public string FrontAttachmentPath { get; set; }
            public string DefectPath { get; set; }
            public string TopCameraSerial { get; set; }
            public string FrontCameraSerial { get; set; }
            public string TurntablePort { get; set; }
            public bool UseRecordedRun { get; set; }
            public string RecordedRunPath { get; set; }
            public string LastPartID { get; set; }

            public static InitializationSettings Load(string path)
            {
                InitializationSettings settings = new InitializationSettings();
                try
                {
                    if (!File.Exists(path))
                    {
                        return settings;
                    }

                    foreach (string rawLine in File.ReadAllLines(path))
                    {
                        if (string.IsNullOrWhiteSpace(rawLine))
                        {
                            continue;
                        }

                        string line = rawLine.Trim();
                        if (line.StartsWith("#", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        int separatorIndex = line.IndexOf('=');
                        if (separatorIndex < 0)
                        {
                            continue;
                        }

                        string key = line.Substring(0, separatorIndex).Trim();
                        string value = line.Substring(separatorIndex + 1).Trim();

                        switch (key)
                        {
                            case "AttachmentPath":
                                settings.AttachmentPath = value;
                                break;
                            case "DefectPath":
                                settings.DefectPath = value;
                                break;
                            case "TopCameraSerial":
                                settings.TopCameraSerial = value;
                                break;
                            case "FrontCameraSerial":
                                settings.FrontCameraSerial = value;
                                break;
                            case "TurntablePort":
                                settings.TurntablePort = value;
                                break;
                            case "UseRecordedRun":
                                settings.UseRecordedRun = bool.TryParse(value, out bool useRecorded) && useRecorded;
                                break;
                            case "RecordedRunPath":
                                settings.RecordedRunPath = value;
                                break;
                            case "LastPartID":
                                settings.LastPartID = value;
                                break;
                        }
                    }
                }
                catch
                {
                    return new InitializationSettings();
                }

                return settings;
            }

            public void Save(string path)
            {
                string directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                StringBuilder builder = new StringBuilder();
                builder.AppendLine($"AttachmentPath={AttachmentPath ?? string.Empty}");
                builder.AppendLine($"DefectPath={DefectPath ?? string.Empty}");
                builder.AppendLine($"TopCameraSerial={TopCameraSerial ?? string.Empty}");
                builder.AppendLine($"FrontCameraSerial={FrontCameraSerial ?? string.Empty}");
                builder.AppendLine($"TurntablePort={TurntablePort ?? string.Empty}");
                builder.AppendLine($"UseRecordedRun={UseRecordedRun}");
                builder.AppendLine($"RecordedRunPath={RecordedRunPath ?? string.Empty}");
                builder.AppendLine($"LastPartID={LastPartID ?? string.Empty}");
                File.WriteAllText(path, builder.ToString());
            }
        }

        [DataContract]
        private sealed class DefectPolicy
        {
            private static readonly StringComparer ClassComparer = StringComparer.OrdinalIgnoreCase;

            [DataMember(Order = 1)]
            public int Version { get; set; } = 1;

            [DataMember(Order = 2)]
            public DateTime LastUpdatedUtc { get; set; } = DateTime.UtcNow;

            [DataMember(Order = 3)]
            public List<DefectClassRule> ClassRules { get; set; } = new List<DefectClassRule>();

            public static DefectPolicy Create()
            {
                return new DefectPolicy();
            }

            public static DefectPolicy Load(string path)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return Create();
                }

                try
                {
                    using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(DefectPolicy));
                        return serializer.ReadObject(stream) as DefectPolicy ?? Create();
                    }
                }
                catch
                {
                    return Create();
                }
            }

            public void Save(string path)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return;
                }

                string directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                LastUpdatedUtc = DateTime.UtcNow;
                using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(DefectPolicy));
                    serializer.WriteObject(stream, this);
                }
            }

            public bool SyncWithClasses(IEnumerable<string> classes)
            {
                if (classes == null)
                {
                    return false;
                }

                Dictionary<string, string> canonical = new Dictionary<string, string>(ClassComparer);
                foreach (string raw in classes)
                {
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        continue;
                    }

                    string trimmed = raw.Trim();
                    if (!canonical.ContainsKey(trimmed))
                    {
                        canonical[trimmed] = trimmed;
                    }
                }

                bool changed = false;

                foreach (DefectClassRule rule in ClassRules)
                {
                    string ruleName = (rule.ClassName ?? string.Empty).Trim();
                    if (canonical.TryGetValue(ruleName, out string canonicalName))
                    {
                        if (!string.Equals(rule.ClassName, canonicalName, StringComparison.Ordinal))
                        {
                            rule.ClassName = canonicalName;
                            changed = true;
                        }

                        if (rule.IsStale)
                        {
                            rule.IsStale = false;
                            changed = true;
                        }
                    }
                    else
                    {
                        if (!rule.IsStale)
                        {
                            rule.IsStale = true;
                            changed = true;
                        }
                    }
            }

            foreach (string canonicalName in canonical.Values)
            {
                if (!ClassRules.Any(rule => ClassComparer.Equals(rule.ClassName, canonicalName)))
                {
                    ClassRules.Add(DefectClassRule.Create(canonicalName));
                    changed = true;
                }
            }

            List<DefectClassRule> staleRules = ClassRules.Where(rule => rule.IsStale).ToList();
            if (staleRules.Count > 0)
            {
                foreach (DefectClassRule stale in staleRules)
                {
                    ClassRules.Remove(stale);
                }
                changed = true;
            }

            if (ClassRules.Count > 1)
            {
                List<string> before = ClassRules.Select(rule => rule.ClassName).ToList();
                ClassRules.Sort((left, right) => ClassComparer.Compare(left.ClassName, right.ClassName));
                if (!before.SequenceEqual(ClassRules.Select(rule => rule.ClassName), ClassComparer))
                    {
                        changed = true;
                    }
                }

                return changed;
            }

            public DefectClassRule GetRule(string className)
            {
                if (string.IsNullOrWhiteSpace(className))
                {
                    return null;
                }

                return ClassRules.FirstOrDefault(rule => ClassComparer.Equals(rule.ClassName, className));
            }
        }

        [DataContract]
        private sealed class DefectClassRule
        {
            [DataMember(Order = 1)]
            public string ClassName { get; set; } = string.Empty;

            [DataMember(Order = 2)]
            public bool Enabled { get; set; } = true;

            [DataMember(Order = 3)]
            public bool FailOnMatch { get; set; } = true;

            [DataMember(Order = 4, EmitDefaultValue = false)]
            public double? MinConfidence { get; set; }

            [DataMember(Order = 5, EmitDefaultValue = false)]
            public double? MaxConfidence { get; set; }

            [DataMember(Order = 6, EmitDefaultValue = false)]
            public double? MinArea { get; set; }

            [DataMember(Order = 7, EmitDefaultValue = false)]
            public double? MaxArea { get; set; }

            [DataMember(Order = 8, EmitDefaultValue = false)]
            public int? MinCount { get; set; }

            [DataMember(Order = 9, EmitDefaultValue = false)]
            public int? MaxCount { get; set; }

            [DataMember(Order = 10)]
            public bool IsStale { get; set; }

            public static DefectClassRule Create(string className)
            {
                return new DefectClassRule
                {
                    ClassName = className?.Trim() ?? string.Empty,
                    Enabled = true,
                    FailOnMatch = true,
                    IsStale = false
                };
            }
        }

        private sealed class RunSession
        {
            public RunSession(string root)
            {
                Root = root ?? throw new ArgumentNullException(nameof(root));
                TopFolder = Path.Combine(root, "Top");
                FrontFolder = Path.Combine(root, "Front");
                ResultsFolder = Path.Combine(root, "Results");
            }

            public string Root { get; }
            public string TopFolder { get; }
            public string FrontFolder { get; }
            public string ResultsFolder { get; }
            public string TopImagePath => Path.Combine(TopFolder, "Top.png");
            public string RecordedSourcePath { get; set; }
        }

        private sealed class InitializationDialog : Form
        {
            private readonly Control content;
            private readonly Panel contentHost;
            private readonly Button closeButton;

            public InitializationDialog(Control content)
            {
                this.content = content ?? throw new ArgumentNullException(nameof(content));

                Text = "System Initialization";
                StartPosition = FormStartPosition.CenterParent;
                Size = new Size(1100, 720);
                MinimumSize = new Size(900, 620);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;

                contentHost = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = Color.FromArgb(246, 248, 252)
                };
                Controls.Add(contentHost);

                if (content.AutoSize)
                {
                    content.Dock = DockStyle.Top;
                }
                else
                {
                    content.Dock = DockStyle.Fill;
                }

                contentHost.Controls.Add(content);

                closeButton = new Button
                {
                    Text = "Close",
                    DialogResult = DialogResult.Cancel,
                    Size = new Size(120, 36),
                    Margin = new Padding(0)
                };

                FlowLayoutPanel footer = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(16, 8, 16, 12),
                    Height = 60
                };
                footer.Controls.Add(closeButton);
                Controls.Add(footer);

                CancelButton = closeButton;
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (content != null && !content.IsDisposed)
                {
                    contentHost.Controls.Remove(content);
                }

                base.OnFormClosing(e);
            }
        }

        private void EnsureInitLayout()
        {
            if (initLayout != null && !initLayout.IsDisposed)
            {
                return;
            }

            initLayout = BuildInitializationLayout();
            BindModelContextUI();
            RefreshCameraUiState(topCameraContext);
            RefreshCameraUiState(frontCameraContext);
            RefreshCameraList();
            UpdateInitSummary();
        }

        private void InitializeResponsiveLayout()
        {
            if (responsiveLayoutInitialized)
            {
                ApplyResponsiveLayout(force: true);
                return;
            }

            responsiveLayoutInitialized = true;
            Resize -= DemoApp_Resize;
            Resize += DemoApp_Resize;
            ApplyResponsiveLayout(force: true);
        }

        private void DemoApp_Resize(object sender, EventArgs e)
        {
            ApplyResponsiveLayout();
        }

        private void ApplyResponsiveLayout(bool force = false)
        {
            if (!responsiveLayoutInitialized && !force)
            {
                return;
            }

            float scaleFactor = DeviceDpi / 96f;
            bool shouldCompact = ClientSize.Width < CompactLayoutWidthThreshold || scaleFactor >= CompactLayoutDpiThreshold;

            if (!force && shouldCompact == usingCompactLayout)
            {
                return;
            }

            usingCompactLayout = shouldCompact;

            tableLayoutPanelMain.SuspendLayout();
            panelStepsHost.SuspendLayout();
            try
            {
                tableLayoutPanelMain.Controls.Clear();
                tableLayoutPanelMain.ColumnStyles.Clear();
                tableLayoutPanelMain.RowStyles.Clear();
                panelStepsHost.Controls.Remove(tableLayoutSteps);
                panelStepsHost.Controls.Clear();

                tableLayoutSteps.Dock = DockStyle.Top;
                tableLayoutSteps.Margin = new Padding(0);
                panelStepsHost.Controls.Add(tableLayoutSteps);
                panelStepsHost.Dock = DockStyle.Fill;
                panelStepsHost.Padding = new Padding(0);
                panelStepsHost.AutoScroll = true;

                Control primaryHost = null;
                if (leftTabs != null && !leftTabs.IsDisposed)
                {
                    primaryHost = leftTabs;
                    leftTabs.Dock = DockStyle.Fill;
                    leftTabs.Margin = new Padding(0);
                }
                else if (panelStepsHost != null)
                {
                    primaryHost = panelStepsHost;
                }
                else
                {
                    primaryHost = tableLayoutSteps;
                }

                if (shouldCompact)
                {
                    tableLayoutPanelMain.ColumnCount = 1;
                    tableLayoutPanelMain.RowCount = 2;
                    tableLayoutPanelMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                    tableLayoutPanelMain.RowStyles.Add(new RowStyle(SizeType.Percent, 52F));
                    tableLayoutPanelMain.RowStyles.Add(new RowStyle(SizeType.Percent, 48F));

                    if (primaryHost != null)
                    {
                        primaryHost.Margin = new Padding(0, 0, 0, ScaleSize(12));
                    }
                    tableLayoutImages.Dock = DockStyle.Fill;
                    tableLayoutImages.Margin = new Padding(0);

                    tableLayoutPanelMain.Controls.Add(primaryHost, 0, 0);
                    tableLayoutPanelMain.Controls.Add(tableLayoutImages, 0, 1);
                }
                else
                {
                    tableLayoutPanelMain.ColumnCount = 2;
                    tableLayoutPanelMain.RowCount = 1;
                    tableLayoutPanelMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
                    tableLayoutPanelMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62F));
                    tableLayoutPanelMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                    if (primaryHost != null)
                    {
                        primaryHost.Margin = new Padding(0, 0, ScaleSize(12), 0);
                    }
                    tableLayoutImages.Dock = DockStyle.Fill;
                    tableLayoutImages.Margin = new Padding(0);

                    tableLayoutPanelMain.Controls.Add(primaryHost, 0, 0);
                    tableLayoutPanelMain.Controls.Add(tableLayoutImages, 1, 0);
                }

                tableLayoutSteps.PerformLayout();
                Size preferred = tableLayoutSteps.PreferredSize;
                panelStepsHost.AutoScrollMinSize = new Size(0, preferred.Height + 16);
            }
            finally
            {
                panelStepsHost.ResumeLayout(true);
                tableLayoutPanelMain.ResumeLayout(true);
            }
        }

        private void AdjustRecordedRunLayout()
        {
            if (recordedRunPanel == null || TB_RecordedRunPath == null || BT_SelectRecordedRun == null || CHK_UseRecordedRun == null)
            {
                return;
            }

            int availableWidth = recordedRunPanel.ClientSize.Width
                - CHK_UseRecordedRun.Width
                - BT_SelectRecordedRun.Width
                - ScaleSize(32);

            availableWidth = Math.Max(ScaleSize(240), availableWidth);

            TB_RecordedRunPath.Width = availableWidth;
            TB_RecordedRunPath.Height = ScaleSize(28);
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

        private void InitializeCycleTimeLabel()
        {
            if (groupStep4 == null || BT_Detect == null)
            {
                return;
            }

            // Increase minimum height of groupStep4
            groupStep4.MinimumSize = new Size(0, ScaleSize(130));

            // Remove BT_Detect temporarily to reorder controls properly
            groupStep4.Controls.Remove(BT_Detect);

            // Create cycle time label (at bottom - add first)
            LBL_CycleTime = new Label
            {
                Dock = DockStyle.Bottom,
                Text = "Cycle: --",
                TextAlign = ContentAlignment.MiddleCenter,
                Height = ScaleSize(28),
                ForeColor = Color.FromArgb(94, 102, 112),
                Font = ScaleFont("Segoe UI", 9F, FontStyle.Bold),
                Padding = new Padding(0, ScaleSize(6), 0, ScaleSize(6))
            };
            groupStep4.Controls.Add(LBL_CycleTime);

            // Create Part ID input panel (at top - add second)
            Panel partIDPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = ScaleSize(36),
                Padding = new Padding(ScaleSize(4), ScaleSize(6), ScaleSize(4), ScaleSize(2))
            };

            LBL_PartID = new Label
            {
                Text = "Part ID:",
                Dock = DockStyle.Left,
                Width = ScaleSize(65),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = ScaleFont("Segoe UI", 9F),
                ForeColor = Color.FromArgb(60, 72, 88)
            };

            TB_PartID = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = ScaleFont("Segoe UI", 9.5F),
                MaxLength = 50
            };

            TB_PartID.TextChanged += (s, e) =>
            {
                currentPartID = TB_PartID.Text?.Trim();
                if (initSettings != null)
                {
                    initSettings.LastPartID = currentPartID;
                    SaveInitializationSettings();
                }
            };

            partIDPanel.Controls.Add(TB_PartID);
            partIDPanel.Controls.Add(LBL_PartID);
            groupStep4.Controls.Add(partIDPanel);

            // Re-add button with Fill dock (add last so it fills remaining space)
            BT_Detect.Dock = DockStyle.Fill;
            BT_Detect.MinimumSize = new Size(0, ScaleSize(44));
            groupStep4.Controls.Add(BT_Detect);
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
                groupBox.Font = ScaleFont("Segoe UI Semibold", 9.5F);
                groupBox.Padding = new Padding(ScaleSize(12), ScaleSize(24), ScaleSize(12), ScaleSize(12));
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

            container.Margin = new Padding(ScaleSize(8));
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

        private sealed class IconButtonMetadata
        {
            public IconButtonMetadata(string glyph, string label, int minWidth, float fontSize, Padding? margin)
            {
                Glyph = glyph ?? string.Empty;
                Label = label ?? string.Empty;
                MinWidth = minWidth;
                FontSize = fontSize;
                CustomMargin = margin;
            }

            public string Glyph { get; }
            public string Label { get; }
            public int MinWidth { get; }
            public float FontSize { get; }
            public Padding? CustomMargin { get; }
        }

        private void ConfigureIconButton(Button button, string glyph, string tooltip, int minWidth = 40, float fontSize = 16f, Padding? customMargin = null)
        {
            if (button == null)
            {
                return;
            }

            string label = button.Text;
            IconButtonMetadata metadata = new IconButtonMetadata(glyph, label, minWidth, fontSize, customMargin);
            button.Tag = metadata;
            button.UseMnemonic = false;
            button.AutoSize = false;
            ApplyIconButtonLayout(button, metadata);

            if (toolTip != null && !string.IsNullOrWhiteSpace(tooltip))
            {
                toolTip.SetToolTip(button, tooltip);
            }
        }

        private void ApplyIconButtonLayout(Button button, IconButtonMetadata metadata)
        {
            if (button == null || metadata == null)
            {
                return;
            }

            button.Image = GetGlyphImage(metadata.Glyph, Color.White, metadata.FontSize);
            button.Text = metadata.Label;
            button.ImageAlign = ContentAlignment.MiddleLeft;
            button.TextAlign = ContentAlignment.MiddleLeft;
            button.TextImageRelation = TextImageRelation.ImageBeforeText;
            button.Font = ScaleFont("Segoe UI Semibold", 9.5F);
            button.Padding = new Padding(ScaleSize(10), 0, ScaleSize(12), 0);
            button.Margin = metadata.CustomMargin ?? new Padding(ScaleSize(6), ScaleSize(4), ScaleSize(6), ScaleSize(4));
            button.MinimumSize = new Size(ScaleSize(metadata.MinWidth), ScaleSize(36));
            button.Height = ScaleSize(36);
        }

        private void StylePrimaryButton(Button button, Color accent, Color hoverAccent)
        {
            if (button == null)
            {
                return;
            }

            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = accent;
            button.ForeColor = Color.White;
            button.Font = ScaleFont("Segoe UI Semibold", 9.5F);
            button.Height = ScaleSize(38);
            button.Margin = new Padding(ScaleSize(8));
            button.Cursor = Cursors.Hand;

            button.MouseEnter += (s, e) => button.BackColor = hoverAccent;
            button.MouseLeave += (s, e) => button.BackColor = accent;

            if (button.Tag is IconButtonMetadata iconMeta)
            {
                ApplyIconButtonLayout(button, iconMeta);
            }
        }

        private void StyleDangerButton(Button button, Color accent, Color hoverAccent)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = accent;
            button.ForeColor = Color.White;
            button.Font = ScaleFont("Segoe UI", 9F);
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

            label.AutoSize = true;
            label.Dock = DockStyle.Top;
            label.BorderStyle = BorderStyle.None;
            label.Padding = new Padding(ScaleSize(16), ScaleSize(10), ScaleSize(16), ScaleSize(10));
            label.Margin = new Padding(ScaleSize(8), ScaleSize(8), ScaleSize(8), 0);
            label.Font = ScaleFont("Segoe UI Semibold", 9.5F);
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
            grid.ColumnHeadersDefaultCellStyle.Font = ScaleFont("Segoe UI Semibold", 9F);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            grid.ColumnHeadersHeight = ScaleSize(34);

            grid.DefaultCellStyle.BackColor = surface;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(223, 230, 253);
            grid.DefaultCellStyle.SelectionForeColor = accent;
            grid.DefaultCellStyle.ForeColor = textColor;
            grid.DefaultCellStyle.Font = ScaleFont("Segoe UI", 9F);
            grid.RowTemplate.Height = ScaleSize(36);

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
                ApplyStatusToAll(StatusPassNoRulesMessage, statusPassBackground, statusPassForeground);
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
                ApplyStatusToAll(StatusPassNoRulesMessage, statusPassBackground, statusPassForeground);
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

        private Image GetGlyphImage(string glyph, Color color, float fontSize)
        {
            if (string.IsNullOrWhiteSpace(glyph))
            {
                return null;
            }

            int iconSize = Math.Max(ScaleSize(16), (int)Math.Round(fontSize * CurrentDpiScale));
            string cacheKey = $"{glyph}-{color.ToArgb()}-{iconSize}";
            if (glyphImageCache.TryGetValue(cacheKey, out Image cached))
            {
                return cached;
            }

            Bitmap bitmap = new Bitmap(iconSize, iconSize);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                using (Font font = new Font(IconFontFamily, fontSize, FontStyle.Regular, GraphicsUnit.Point))
                using (Brush brush = new SolidBrush(color))
                {
                    RectangleF bounds = new RectangleF(0, 0, bitmap.Width, bitmap.Height);
                    StringFormat format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    graphics.DrawString(glyph, font, brush, bounds, format);
                }
            }

            glyphImageCache[cacheKey] = bitmap;
            return bitmap;
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
            if (CB_Value == null || CB_Value.IsDisposed)
            {
                return;
            }

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

            bool beganUpdate = false;
            try
            {
                if (CB_Value.IsHandleCreated)
                {
                    CB_Value.BeginUpdate();
                    beganUpdate = true;
                }

                CB_Value.AutoCompleteMode = AutoCompleteMode.None;
                CB_Value.AutoCompleteSource = AutoCompleteSource.None;
                CB_Value.Items.Clear();

                foreach (string suggestion in suggestions)
                {
                    CB_Value.Items.Add(suggestion);
                }

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
            }
            finally
            {
                if (beganUpdate && CB_Value.IsHandleCreated)
                {
                    CB_Value.EndUpdate();
                }

                suppressLogicEvents = previousSuppress;
            }

        }

        private void RefreshClassRuleCardsUI()
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new MethodInvoker(RefreshClassRuleCardsUI));
                }
                catch (ObjectDisposedException)
                {
                }
                return;
            }

            if (flowClassRules == null || flowClassRules.IsDisposed)
            {
                return;
            }

            flowClassRules.SuspendLayout();
            flowClassRules.Controls.Clear();
            classRuleCards.Clear();

            if (defectPolicy?.ClassRules != null)
            {
                IEnumerable<DefectClassRule> ordered = defectPolicy.ClassRules
                    .OrderBy(rule => rule.IsStale)
                    .ThenBy(rule => rule.ClassName, StringComparer.OrdinalIgnoreCase);

                foreach (DefectClassRule rule in ordered)
                {
                    ClassRuleCard card = CreateClassRuleCard(rule);
                    classRuleCards[rule.ClassName] = card;
                    flowClassRules.Controls.Add(card.Container);
                    ApplyRuleToCard(card);
                }
            }

            flowClassRules.ResumeLayout(true);
        }

        private ClassRuleCard CreateClassRuleCard(DefectClassRule rule)
        {
            ClassRuleCard card = new ClassRuleCard(rule);

            GroupBox container = new GroupBox
            {
                Text = string.IsNullOrWhiteSpace(rule.ClassName) ? "Unnamed Class" : rule.ClassName,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(ScaleSize(12), ScaleSize(24), ScaleSize(12), ScaleSize(12)),
                Margin = new Padding(ScaleSize(8)),
                MinimumSize = new Size(ScaleSize(420), ScaleSize(0)),
                Font = ScaleFont("Segoe UI Semibold", 9.25F)
            };
            card.Container = container;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                ColumnCount = 3,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            container.Controls.Add(layout);

            card.EnableCheck = new CheckBox
            {
                Text = "Enable rule for this class",
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, ScaleSize(6)),
                Font = ScaleFont("Segoe UI", 9F)
            };
            card.EnableCheck.CheckedChanged += OnClassRuleEnabledChanged;
            card.EnableCheck.Tag = card;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(card.EnableCheck, 0, 0);
            layout.SetColumnSpan(card.EnableCheck, 3);

            Label outcomeLabel = new Label
            {
                Text = "When detections meet these thresholds",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, ScaleSize(4), ScaleSize(6)),
                Font = ScaleFont("Segoe UI", 9F, FontStyle.Regular)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(outcomeLabel, 0, 1);

            card.OutcomeCombo = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, ScaleSize(4), ScaleSize(6)),
                Font = ScaleFont("Segoe UI", 9F)
            };
            card.OutcomeCombo.Items.Add(new ComboOption<bool>("Fail the inspection", true));
            card.OutcomeCombo.Items.Add(new ComboOption<bool>("Pass only when thresholds remain satisfied", false));
            card.OutcomeCombo.SelectedIndexChanged += OnClassRuleOutcomeChanged;
            card.OutcomeCombo.Tag = card;
            layout.Controls.Add(card.OutcomeCombo, 1, 1);
            layout.SetColumnSpan(card.OutcomeCombo, 2);

            int rowIndex = 2;

            AddThresholdRow(card, layout, ref rowIndex, "Minimum confidence", "score", 3, 0m, 1m, 0.01m,
                r => r.MinConfidence, (r, v) => r.MinConfidence = v);
            AddThresholdRow(card, layout, ref rowIndex, "Maximum confidence", "score", 3, 0m, 1m, 0.01m,
                r => r.MaxConfidence, (r, v) => r.MaxConfidence = v);
            AddThresholdRow(card, layout, ref rowIndex, "Minimum area", "px^2", 0, 0m, 10000000m, 10m,
                r => r.MinArea, (r, v) => r.MinArea = v);
            AddThresholdRow(card, layout, ref rowIndex, "Maximum area", "px^2", 0, 0m, 10000000m, 10m,
                r => r.MaxArea, (r, v) => r.MaxArea = v);
            AddThresholdRow(card, layout, ref rowIndex, "Minimum detections", "count", 0, 0m, 10000m, 1m,
                r => r.MinCount.HasValue ? (double?)r.MinCount.Value : null,
                (r, v) => r.MinCount = v.HasValue ? (int?)Math.Round(v.Value) : null);
            AddThresholdRow(card, layout, ref rowIndex, "Maximum detections", "count", 0, 0m, 10000m, 1m,
                r => r.MaxCount.HasValue ? (double?)r.MaxCount.Value : null,
                (r, v) => r.MaxCount = v.HasValue ? (int?)Math.Round(v.Value) : null);

            card.StatusLabel = new Label
            {
                Text = "Class not present in the currently loaded project.",
                ForeColor = Color.FromArgb(178, 34, 34),
                Dock = DockStyle.Fill,
                Visible = rule.IsStale,
                Margin = new Padding(0, ScaleSize(4), 0, 0),
                Font = ScaleFont("Segoe UI", 8.5F, FontStyle.Italic)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (layout.RowCount <= rowIndex)
            {
                layout.RowCount = rowIndex + 1;
            }
            layout.Controls.Add(card.StatusLabel, 0, rowIndex);
            layout.SetColumnSpan(card.StatusLabel, 3);
            rowIndex++;
            layout.RowCount = Math.Max(layout.RowCount, rowIndex);

            return card;
        }

        private ThresholdPair AddThresholdRow(
            ClassRuleCard card,
            TableLayoutPanel layout,
            ref int rowIndex,
            string labelText,
            string unitText,
            int decimalPlaces,
            decimal minimum,
            decimal maximum,
            decimal increment,
            Func<DefectClassRule, double?> getter,
            Action<DefectClassRule, double?> setter)
        {
            ThresholdPair pair = new ThresholdPair(card, labelText, getter, setter);

            CheckBox toggle = new CheckBox
            {
                Text = labelText,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, ScaleSize(4), ScaleSize(4)),
                Font = ScaleFont("Segoe UI", 9F)
            };
            toggle.CheckedChanged += OnThresholdToggleChanged;
            toggle.Tag = pair;
            pair.Toggle = toggle;

            NumericUpDown numeric = new NumericUpDown
            {
                DecimalPlaces = decimalPlaces,
                Minimum = minimum,
                Maximum = maximum,
                Increment = increment,
                Dock = DockStyle.Fill,
                Enabled = false,
                Margin = new Padding(0, 0, ScaleSize(4), ScaleSize(4)),
                ThousandsSeparator = decimalPlaces == 0,
                Font = ScaleFont("Segoe UI", 9F)
            };
            numeric.ValueChanged += OnThresholdValueChanged;
            numeric.Tag = pair;
            pair.Numeric = numeric;

            Label unitLabel = new Label
            {
                Text = unitText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, ScaleSize(4)),
                Font = ScaleFont("Segoe UI", 8.5F, FontStyle.Italic),
                Enabled = false
            };
            pair.UnitLabel = unitLabel;

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (layout.RowCount <= rowIndex)
            {
                layout.RowCount = rowIndex + 1;
            }

            layout.Controls.Add(toggle, 0, rowIndex);
            layout.Controls.Add(numeric, 1, rowIndex);
            layout.Controls.Add(unitLabel, 2, rowIndex);
            rowIndex++;
            layout.RowCount = Math.Max(layout.RowCount, rowIndex);

            card.Thresholds.Add(pair);
            return pair;
        }

        private void ApplyRuleToCard(ClassRuleCard card)
        {
            if (card == null)
            {
                return;
            }

            card.Updating = true;

            card.Container.Text = string.IsNullOrWhiteSpace(card.Rule.ClassName) ? "Unnamed Class" : card.Rule.ClassName;
            card.EnableCheck.Checked = card.Rule.Enabled;
            card.EnableCheck.Enabled = !card.Rule.IsStale;

            bool interactable = card.Rule.Enabled && !card.Rule.IsStale;

            int outcomeIndex = card.Rule.FailOnMatch ? 0 : 1;
            if (card.OutcomeCombo.Items.Count > outcomeIndex)
            {
                card.OutcomeCombo.SelectedIndex = outcomeIndex;
            }
            card.OutcomeCombo.Enabled = interactable;

            foreach (ThresholdPair pair in card.Thresholds)
            {
                double? value = pair.Getter(card.Rule);
                bool toggleEnabled = !card.Rule.IsStale && card.Rule.Enabled;
                pair.Toggle.Enabled = toggleEnabled;
                pair.Toggle.Checked = value.HasValue;

                bool numericEnabled = value.HasValue && interactable;
                pair.Numeric.Enabled = numericEnabled;
                pair.UnitLabel.Enabled = numericEnabled;

                if (value.HasValue)
                {
                    pair.Numeric.Value = ClampToNumeric(pair.Numeric, value.Value);
                }
            }

            if (card.StatusLabel != null)
            {
                card.StatusLabel.Visible = card.Rule.IsStale;
            }

            card.Updating = false;
        }

        private static decimal ClampToNumeric(NumericUpDown numeric, double value)
        {
            decimal dec;
            try
            {
                dec = Convert.ToDecimal(value);
            }
            catch
            {
                dec = numeric.Minimum;
            }

            if (dec < numeric.Minimum)
            {
                dec = numeric.Minimum;
            }
            if (dec > numeric.Maximum)
            {
                dec = numeric.Maximum;
            }

            return dec;
        }

        private void OnClassRuleEnabledChanged(object sender, EventArgs e)
        {
            if (sender is CheckBox check && check.Tag is ClassRuleCard card)
            {
                if (card.Updating || card.Rule.IsStale)
                {
                    return;
                }

                card.Updating = true;
                card.Rule.Enabled = check.Checked;
                card.Updating = false;
                ApplyRuleToCard(card);
                SaveDefectPolicyToDisk();
                outToLog($"[Policy] {(card.Rule.Enabled ? "Enabled" : "Disabled")} class rule '{card.Rule.ClassName}'.", LogStatus.Info);
            }
        }

        private void OnClassRuleOutcomeChanged(object sender, EventArgs e)
        {
            if (sender is ComboBox combo && combo.Tag is ClassRuleCard card)
            {
                if (card.Updating || card.Rule.IsStale)
                {
                    return;
                }

                if (combo.SelectedItem is ComboOption<bool> option)
                {
                    card.Updating = true;
                    card.Rule.FailOnMatch = option.Value;
                    card.Updating = false;
                    SaveDefectPolicyToDisk();
                }
            }
        }

        private void OnThresholdToggleChanged(object sender, EventArgs e)
        {
            if (sender is CheckBox toggle && toggle.Tag is ThresholdPair pair)
            {
                ClassRuleCard card = pair.Card;
                if (card.Updating || card.Rule.IsStale)
                {
                    return;
                }

                card.Updating = true;
                bool enabled = toggle.Checked && card.Rule.Enabled && !card.Rule.IsStale;
                pair.Numeric.Enabled = enabled;
                pair.UnitLabel.Enabled = enabled;
                if (!card.Rule.Enabled && toggle.Checked)
                {
                    // revert to stored value if rule is disabled
                    toggle.Checked = pair.Getter(card.Rule).HasValue;
                }
                else
                {
                    pair.Setter(card.Rule, enabled ? (double)pair.Numeric.Value : (double?)null);
                }
                card.Updating = false;
                SaveDefectPolicyToDisk();
            }
        }

        private void OnThresholdValueChanged(object sender, EventArgs e)
        {
            if (sender is NumericUpDown numeric && numeric.Tag is ThresholdPair pair)
            {
                ClassRuleCard card = pair.Card;
                if (card.Updating || card.Rule.IsStale || !pair.Toggle.Checked || !card.Rule.Enabled)
                {
                    return;
                }

                card.Updating = true;
                pair.Setter(card.Rule, (double)numeric.Value);
                card.Updating = false;
                SaveDefectPolicyToDisk();
            }
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

        private sealed class ClassRuleCard
        {
            public ClassRuleCard(DefectClassRule rule)
            {
                Rule = rule ?? throw new ArgumentNullException(nameof(rule));
                Thresholds = new List<ThresholdPair>();
            }

            public DefectClassRule Rule { get; }
            public GroupBox Container { get; set; }
            public CheckBox EnableCheck { get; set; }
            public ComboBox OutcomeCombo { get; set; }
            public Label StatusLabel { get; set; }
            public List<ThresholdPair> Thresholds { get; }
            public bool Updating { get; set; }
        }

        private sealed class ThresholdPair
        {
            public ThresholdPair(ClassRuleCard card, string name, Func<DefectClassRule, double?> getter, Action<DefectClassRule, double?> setter)
            {
                Card = card ?? throw new ArgumentNullException(nameof(card));
                PropertyName = name ?? string.Empty;
                Getter = getter ?? throw new ArgumentNullException(nameof(getter));
                Setter = setter ?? throw new ArgumentNullException(nameof(setter));
            }

            public ClassRuleCard Card { get; }
            public string PropertyName { get; }
            public CheckBox Toggle { get; set; }
            public NumericUpDown Numeric { get; set; }
            public Label UnitLabel { get; set; }
            public Func<DefectClassRule, double?> Getter { get; }
            public Action<DefectClassRule, double?> Setter { get; }
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
        private sealed class ConsoleRedirectWriter : TextWriter
        {
            private readonly TextWriter originalWriter;
            private readonly Action<LogStatus, string> lineHandler;
            private readonly LogStatus defaultStatus;
            private readonly StringBuilder buffer = new StringBuilder();
            private readonly object syncRoot = new object();

            public ConsoleRedirectWriter(TextWriter originalWriter, Action<LogStatus, string> lineHandler, LogStatus defaultStatus)
            {
                this.originalWriter = originalWriter ?? TextWriter.Null;
                this.lineHandler = lineHandler;
                this.defaultStatus = defaultStatus;
            }

            public override Encoding Encoding => originalWriter.Encoding ?? Encoding.UTF8;

            public override void Write(char value)
            {
                originalWriter.Write(value);
                AppendChar(value);
            }

            public override void Write(string value)
            {
                originalWriter.Write(value);
                if (value == null)
                {
                    return;
                }

                foreach (char c in value)
                {
                    AppendChar(c);
                }
            }

            public override void WriteLine(string value)
            {
                originalWriter.WriteLine(value);
                if (value != null)
                {
                    foreach (char c in value)
                    {
                        AppendChar(c);
                    }
                }
                AppendChar('\n');
            }

            public override void Flush()
            {
                originalWriter.Flush();
                FlushPending();
            }

            public void FlushPending()
            {
                string line = null;
                lock (syncRoot)
                {
                    if (buffer.Length > 0)
                    {
                        line = buffer.ToString();
                        buffer.Clear();
                    }
                }

                if (!string.IsNullOrEmpty(line))
                {
                    Emit(line);
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    FlushPending();
                }

                base.Dispose(disposing);
            }

            private void AppendChar(char value)
            {
                string line = null;
                lock (syncRoot)
                {
                    if (value == '\r')
                    {
                        return;
                    }

                    if (value == '\n')
                    {
                        if (buffer.Length > 0)
                        {
                            line = buffer.ToString();
                            buffer.Clear();
                        }
                    }
                    else
                    {
                        buffer.Append(value);
                    }
                }

                if (!string.IsNullOrEmpty(line))
                {
                    Emit(line);
                }
            }

            private void Emit(string line)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    return;
                }

                lineHandler?.Invoke(defaultStatus, line.Trim());
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
    public FrontCaptureTicket(int sequence, string imagePath, string originalImagePath, double angleDegrees, DateTime capturedAt)
    {
        Sequence = sequence;
        ImagePath = imagePath;
        OriginalImagePath = originalImagePath;
        AngleDegrees = angleDegrees;
        CapturedAt = capturedAt;
    }

    public int Sequence { get; }
    public string ImagePath { get; }
    public string OriginalImagePath { get; }
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

// JSON export data structures
[Serializable]
internal class TopAttachmentData
{
    [JsonProperty("sequence")]
    public int Sequence { get; set; }

    [JsonProperty("center")]
    public PointData Center { get; set; }

    [JsonProperty("angle_degrees")]
    public double AngleDegrees { get; set; }

    [JsonProperty("bounding_box")]
    public BoundingBoxData BoundingBox { get; set; }

    [JsonProperty("class_name")]
    public string ClassName { get; set; }

    [JsonProperty("confidence")]
    public float Confidence { get; set; }
}

[Serializable]
internal class PointData
{
    [JsonProperty("x")]
    public float X { get; set; }

    [JsonProperty("y")]
    public float Y { get; set; }
}

[Serializable]
internal class BoundingBoxData
{
    [JsonProperty("x")]
    public int X { get; set; }

    [JsonProperty("y")]
    public int Y { get; set; }

    [JsonProperty("width")]
    public int Width { get; set; }

    [JsonProperty("height")]
    public int Height { get; set; }
}

[Serializable]
internal class TopAttachmentsJson
{
    [JsonProperty("image_path")]
    public string ImagePath { get; set; }

    [JsonProperty("image_center")]
    public PointData ImageCenter { get; set; }

    [JsonProperty("attachments")]
    public List<TopAttachmentData> Attachments { get; set; }

    [JsonProperty("timestamp")]
    public string Timestamp { get; set; }
}

[Serializable]
internal class FrontAttachmentData
{
    [JsonProperty("center")]
    public PointData Center { get; set; }

    [JsonProperty("bounding_box")]
    public BoundingBoxData BoundingBox { get; set; }

    [JsonProperty("class_name")]
    public string ClassName { get; set; }

    [JsonProperty("confidence")]
    public float Confidence { get; set; }

    [JsonProperty("distance_from_target_center")]
    public int DistanceFromTargetCenter { get; set; }

    [JsonProperty("is_selected_for_crop")]
    public bool IsSelectedForCrop { get; set; }
}

[Serializable]
internal class FrontSequenceData
{
    [JsonProperty("sequence")]
    public int Sequence { get; set; }

    [JsonProperty("image_path")]
    public string ImagePath { get; set; }

    [JsonProperty("angle_degrees")]
    public double AngleDegrees { get; set; }

    [JsonProperty("attachments_detected")]
    public List<FrontAttachmentData> AttachmentsDetected { get; set; }

    [JsonProperty("selected_attachment_center")]
    public PointData SelectedAttachmentCenter { get; set; }

    [JsonProperty("crop_image_path")]
    public string CropImagePath { get; set; }
}

[Serializable]
internal class FrontAttachmentsJson
{
    [JsonProperty("sequences")]
    public List<FrontSequenceData> Sequences { get; set; }

    [JsonProperty("timestamp")]
    public string Timestamp { get; set; }
}
}




















