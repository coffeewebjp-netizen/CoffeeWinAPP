using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices; // Windows API利用に必要
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using GregsStack.InputSimulatorStandard.Native;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using static CoffeeAutoButton.NativeMethods;

namespace CoffeeAutoButton
{
    public partial class MainWindow : Window
    {
        private readonly KeyboardInputService _keyboardInputService = new KeyboardInputService();
        private readonly MouseClickService _mouseClickService = new MouseClickService();
        private readonly BrowserDirectClickService _browserDirectClickService = new BrowserDirectClickService();
        private readonly DispatcherTimer _timer = new DispatcherTimer();
        private readonly DispatcherTimer _statusTimer = new DispatcherTimer();

        // 実行用変数
        private VirtualKeyCode _targetKey = VirtualKeyCode.NONAME;
        private System.Windows.Point _targetPoint;
        private System.Windows.Point _targetClientPoint;
        private bool _hasTargetClientPoint;
        private IntPtr _targetWindowHandle = IntPtr.Zero;
        private string _targetWindowTitle = string.Empty;
        private string _targetWindowClass = string.Empty;
        private int _targetProcessId;
        private string _targetProcessName = string.Empty;
        private string _targetElevationStatus = "権限不明";
        private bool _targetMayRequireElevation;
        private bool _isCurrentProcessElevated;
        private RECT _targetWindowRect;
        private bool _hasTargetWindowRect;
        private DateTime _startTime;
        private DateTime _pausedAt;
        private int _durationSeconds;
        private bool _isPaused;
        private bool _isClickInProgress;
        private bool _isKeyInProgress;
        private bool _isApplyingSettings;
        private bool _isSettingsReady;
        private uint _stopHotkeyModifiers = HotkeyDefaults.StopModifiers;
        private uint _stopHotkeyKey = HotkeyDefaults.StopKey;
        private uint _pauseHotkeyModifiers = HotkeyDefaults.PauseModifiers;
        private uint _pauseHotkeyKey = HotkeyDefaults.PauseKey;
        private HwndSource _windowSource;
        private CancellationTokenSource _startDelayCts;
        private Process _dedicatedBrowserProcess;
        private BrowserClickTarget _browserClickTarget;

        // プリセット用変数
        private List<Preset> _presets = new List<Preset>();

        private const int StopHotkeyId = 1;
        private const int PauseHotkeyId = 2;
        private const int DedicatedBrowserDebugPort = 9223;

        public MainWindow()
        {
            InitializeComponent();
            Title = $"Coffee AutoButton {GetAppVersionLabel()}";
            txtVersion.Text = $"Version {GetAppVersionLabel()}";
            _isCurrentProcessElevated = TryGetProcessElevation(Process.GetCurrentProcess().Id, out var isElevated, out _)
                && isElevated;
            _timer.Tick += Timer_Tick;
            _statusTimer.Interval = TimeSpan.FromMilliseconds(250);
            _statusTimer.Tick += StatusTimer_Tick;
            SourceInitialized += MainWindow_SourceInitialized;
            Closed += MainWindow_Closed;
            
            // プリセット読み込みとUI反映
            _presets = PresetManager.LoadPresets();
            LoadPresetsToUI();
            ApplySettings(AppSettingsManager.Load());
            _isSettingsReady = true;
        }


        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            _windowSource = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
            _windowSource?.AddHook(WndProc);
            RegisterConfiguredHotkeys();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveLastSettings();
            var handle = new WindowInteropHelper(this).Handle;
            if (handle != IntPtr.Zero)
            {
                UnregisterHotKey(handle, StopHotkeyId);
                UnregisterHotKey(handle, PauseHotkeyId);
            }

            _windowSource?.RemoveHook(WndProc);
        }

        private static string GetAppVersionLabel()
        {
            var version = typeof(MainWindow).Assembly.GetName().Version;
            if (version == null)
            {
                return string.Empty;
            }

            return $"v{version.Major}.{version.Minor}.{version.Build}";
        }

    }
}
