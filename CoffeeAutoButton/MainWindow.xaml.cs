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

        private void LoadPresetsToUI()
        {
            cmbPresets.ItemsSource = null;
            cmbPresets.ItemsSource = _presets;
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

        private bool RegisterConfiguredHotkeys()
        {
            var handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero)
            {
                return false;
            }

            UnregisterHotKey(handle, StopHotkeyId);
            UnregisterHotKey(handle, PauseHotkeyId);

            var failedHotkeys = new List<string>();
            if (!RegisterHotKey(handle, StopHotkeyId, _stopHotkeyModifiers, _stopHotkeyKey))
            {
                failedHotkeys.Add($"停止({FormatHotkey(_stopHotkeyModifiers, _stopHotkeyKey)})");
            }

            if (!RegisterHotKey(handle, PauseHotkeyId, _pauseHotkeyModifiers, _pauseHotkeyKey))
            {
                failedHotkeys.Add($"一時停止({FormatHotkey(_pauseHotkeyModifiers, _pauseHotkeyKey)})");
            }

            if (failedHotkeys.Count > 0)
            {
                txtPickStatus.Text = $"{string.Join("、", failedHotkeys)}ホットキーを登録できませんでした";
                return false;
            }

            return true;
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY && wParam.ToInt32() == StopHotkeyId)
            {
                if (_timer.IsEnabled || _startDelayCts != null || _isPaused)
                {
                    StopTimer();
                    txtRunStatus.Text = "ホットキーで停止しました";
                }

                handled = true;
            }
            else if (msg == WM_HOTKEY && wParam.ToInt32() == PauseHotkeyId)
            {
                TogglePauseResume();
                handled = true;
            }

            return IntPtr.Zero;
        }

        private void BtnSavePreset_Click(object sender, RoutedEventArgs e)
        {
            // 入力ダイアログ (VB.NETの機能を使用)
            string name = Interaction.InputBox("プリセット名を入力してください", "設定の保存", "設定1");
            if (string.IsNullOrWhiteSpace(name)) return;

            // 現在の設定を取得
            var preset = new Preset
            {
                Name = name,
                ModeIndex = cmbAction.SelectedIndex,
                ClickMethodIndex = cmbClickMethod.SelectedIndex,
                ClickActionIndex = cmbClickAction.SelectedIndex,
                HoldDurationText = txtHoldDuration.Text,
                AutoFallbackToPhysicalClick = chkAutoFallbackToPhysical.IsChecked == true,
                DedicatedBrowserUrl = txtDedicatedBrowserUrl.Text,
                TargetKey = _targetKey,
                KeyActionIndex = cmbKeyAction.SelectedIndex,
                KeyHoldDurationText = txtKeyHoldDuration.Text,
                KeySequenceText = txtKeySequence.Text,
                KeyModifierCtrl = chkKeyCtrl.IsChecked == true,
                KeyModifierShift = chkKeyShift.IsChecked == true,
                KeyModifierAlt = chkKeyAlt.IsChecked == true,
                KeyModifierWin = chkKeyWin.IsChecked == true,
                TargetPoint = new System.Windows.Point(
                    string.IsNullOrEmpty(txtX.Text) ? 0 : int.Parse(txtX.Text),
                    string.IsNullOrEmpty(txtY.Text) ? 0 : int.Parse(txtY.Text)),
                TargetClientPoint = _targetClientPoint,
                HasTargetClientPoint = _hasTargetClientPoint,
                TargetWindowTitle = _targetWindowTitle,
                TargetWindowClass = _targetWindowClass,
                TargetProcessId = _targetProcessId,
                TargetProcessName = _targetProcessName,
                TargetWindowLeft = _targetWindowRect.Left,
                TargetWindowTop = _targetWindowRect.Top,
                TargetWindowRight = _targetWindowRect.Right,
                TargetWindowBottom = _targetWindowRect.Bottom,
                HasTargetWindowRect = _hasTargetWindowRect,
                IntervalText = txtInterval.Text,
                DurationText = txtDuration.Text,
                StartDelayText = txtStartDelay.Text
            };

            // 同名があれば上書き、なければ追加
            var existing = _presets.Find(p => p.Name == name);
            if (existing != null)
            {
                _presets.Remove(existing);
            }
            _presets.Add(preset);

            PresetManager.SavePresets(_presets);
            LoadPresetsToUI();
            cmbPresets.SelectedItem = preset;
            MessageBox.Show("設定を保存しました。", "完了");
        }

        private void BtnDeletePreset_Click(object sender, RoutedEventArgs e)
        {
            if (cmbPresets.SelectedItem is Preset selected)
            {
                if (MessageBox.Show($"プリセット「{selected.Name}」を削除しますか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    _presets.Remove(selected);
                    PresetManager.SavePresets(_presets);
                    LoadPresetsToUI();
                }
            }
        }

        private void CmbPresets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPresets.SelectedItem is Preset selected)
            {
                _isApplyingSettings = true;
                try
                {
                    // UIに反映
                    cmbAction.SelectedIndex = selected.ModeIndex;
                    txtKeySequence.Text = selected.KeySequenceText ?? string.Empty;
                    if (selected.ModeIndex == 0) // Key
                    {
                        _targetKey = selected.TargetKey;
                        var key = KeyInterop.KeyFromVirtualKey((int)_targetKey);
                        txtKeyDisplay.Text = $"{key} (コード: {_targetKey})";
                        cmbKeyAction.SelectedIndex = selected.KeyActionIndex == 1 ? 1 : 0;
                        txtKeyHoldDuration.Text = string.IsNullOrWhiteSpace(selected.KeyHoldDurationText) ? "500" : selected.KeyHoldDurationText;
                        chkKeyCtrl.IsChecked = selected.KeyModifierCtrl;
                        chkKeyShift.IsChecked = selected.KeyModifierShift;
                        chkKeyAlt.IsChecked = selected.KeyModifierAlt;
                        chkKeyWin.IsChecked = selected.KeyModifierWin;
                    }
                    else // Mouse
                    {
                        _targetPoint = selected.TargetPoint;
                        _targetClientPoint = selected.TargetClientPoint;
                        _hasTargetClientPoint = selected.HasTargetClientPoint;
                        _targetWindowTitle = selected.TargetWindowTitle ?? string.Empty;
                        _targetWindowClass = selected.TargetWindowClass ?? string.Empty;
                        _targetProcessId = selected.TargetProcessId;
                        _targetProcessName = selected.TargetProcessName ?? string.Empty;
                        RefreshTargetElevationStatus();
                        _targetWindowRect = new RECT
                        {
                            Left = selected.TargetWindowLeft,
                            Top = selected.TargetWindowTop,
                            Right = selected.TargetWindowRight,
                            Bottom = selected.TargetWindowBottom
                        };
                        _hasTargetWindowRect = selected.HasTargetWindowRect;
                        cmbClickMethod.SelectedIndex = selected.ClickMethodIndex == 1 ? 1 : 0;
                        cmbClickAction.SelectedIndex = Math.Clamp(selected.ClickActionIndex, 0, 3);
                        txtHoldDuration.Text = string.IsNullOrWhiteSpace(selected.HoldDurationText) ? "500" : selected.HoldDurationText;
                        chkAutoFallbackToPhysical.IsChecked = selected.AutoFallbackToPhysicalClick;
                        txtDedicatedBrowserUrl.Text = selected.DedicatedBrowserUrl ?? string.Empty;
                        txtX.Text = selected.TargetPoint.X.ToString();
                        txtY.Text = selected.TargetPoint.Y.ToString();
                        UpdateTargetInfoText();
                    }
                    txtInterval.Text = selected.IntervalText;
                    txtDuration.Text = selected.DurationText;
                    txtStartDelay.Text = string.IsNullOrWhiteSpace(selected.StartDelayText) ? "0" : selected.StartDelayText;
                }
                finally
                {
                    _isApplyingSettings = false;
                }

                SaveLastSettingsIfReady();
            }
        }

        // --- UI切り替え ---
        private void CmbAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 初期化前の呼び出し防止
            if (pnlKeySettings == null || pnlMouseSettings == null) return;

            if (cmbAction.SelectedIndex == 0) // キー連打モード
            {
                pnlKeySettings.Visibility = Visibility.Visible;
                pnlMouseSettings.Visibility = Visibility.Collapsed;
            }
            else // クリック連打モード
            {
                pnlKeySettings.Visibility = Visibility.Collapsed;
                pnlMouseSettings.Visibility = Visibility.Visible;
            }

            SaveLastSettingsIfReady();
        }

        private void PersistedSetting_Changed(object sender, RoutedEventArgs e)
        {
            SaveLastSettingsIfReady();
        }

        private void PersistedSetting_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SaveLastSettingsIfReady();
        }

        private void PersistedSetting_TextChanged(object sender, TextChangedEventArgs e)
        {
            SaveLastSettingsIfReady();
        }

        private void SaveLastSettingsIfReady()
        {
            if (!_isSettingsReady || _isApplyingSettings)
            {
                return;
            }

            SaveLastSettings();
        }

        private void TxtStopHotkey_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            CaptureHotkey(e, isStopHotkey: true);
        }

        private void TxtPauseHotkey_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            CaptureHotkey(e, isStopHotkey: false);
        }

        private void CaptureHotkey(KeyEventArgs e, bool isStopHotkey)
        {
            e.Handled = true;

            if (_timer.IsEnabled || _startDelayCts != null || _isPaused)
            {
                txtPickStatus.Text = "実行中はホットキーを変更できません";
                return;
            }

            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key == Key.Escape)
            {
                ApplyCapturedHotkey(
                    isStopHotkey,
                    isStopHotkey ? HotkeyDefaults.StopModifiers : HotkeyDefaults.PauseModifiers,
                    isStopHotkey ? HotkeyDefaults.StopKey : HotkeyDefaults.PauseKey);
                return;
            }

            if (IsModifierKey(key))
            {
                return;
            }

            var modifiers = GetHotkeyModifiers(Keyboard.Modifiers);
            if (modifiers == 0)
            {
                txtPickStatus.Text = "ホットキーは Ctrl / Shift / Alt / Win のいずれかを含めてください";
                return;
            }

            var virtualKey = (uint)KeyInterop.VirtualKeyFromKey(key);
            if (virtualKey == 0)
            {
                txtPickStatus.Text = "このキーはホットキーに設定できません";
                return;
            }

            if (IsDuplicateHotkey(isStopHotkey, modifiers, virtualKey))
            {
                txtPickStatus.Text = "停止と一時停止に同じホットキーは設定できません";
                return;
            }

            ApplyCapturedHotkey(isStopHotkey, modifiers, virtualKey);
        }

        private void ApplyCapturedHotkey(bool isStopHotkey, uint modifiers, uint virtualKey)
        {
            var previousStopModifiers = _stopHotkeyModifiers;
            var previousStopKey = _stopHotkeyKey;
            var previousPauseModifiers = _pauseHotkeyModifiers;
            var previousPauseKey = _pauseHotkeyKey;

            if (isStopHotkey)
            {
                _stopHotkeyModifiers = modifiers;
                _stopHotkeyKey = virtualKey;
            }
            else
            {
                _pauseHotkeyModifiers = modifiers;
                _pauseHotkeyKey = virtualKey;
            }

            UpdateHotkeyTextBoxes();

            if (RegisterConfiguredHotkeys())
            {
                txtPickStatus.Text = "ホットキーを更新しました";
                SaveLastSettingsIfReady();
                return;
            }

            var failedMessage = txtPickStatus.Text;
            _stopHotkeyModifiers = previousStopModifiers;
            _stopHotkeyKey = previousStopKey;
            _pauseHotkeyModifiers = previousPauseModifiers;
            _pauseHotkeyKey = previousPauseKey;
            UpdateHotkeyTextBoxes();
            RegisterConfiguredHotkeys();
            txtPickStatus.Text = failedMessage;
        }

        private bool IsDuplicateHotkey(bool isStopHotkey, uint modifiers, uint virtualKey)
        {
            if (isStopHotkey)
            {
                return modifiers == _pauseHotkeyModifiers && virtualKey == _pauseHotkeyKey;
            }

            return modifiers == _stopHotkeyModifiers && virtualKey == _stopHotkeyKey;
        }

        private static bool IsModifierKey(Key key)
        {
            return key == Key.LeftCtrl
                || key == Key.RightCtrl
                || key == Key.LeftShift
                || key == Key.RightShift
                || key == Key.LeftAlt
                || key == Key.RightAlt
                || key == Key.LWin
                || key == Key.RWin
                || key == Key.System;
        }

        private static uint GetHotkeyModifiers(ModifierKeys modifiers)
        {
            uint result = 0;
            if ((modifiers & ModifierKeys.Alt) != 0)
            {
                result |= MOD_ALT;
            }
            if ((modifiers & ModifierKeys.Control) != 0)
            {
                result |= MOD_CONTROL;
            }
            if ((modifiers & ModifierKeys.Shift) != 0)
            {
                result |= MOD_SHIFT;
            }
            if ((modifiers & ModifierKeys.Windows) != 0)
            {
                result |= MOD_WIN;
            }

            return result;
        }

        private void UpdateHotkeyTextBoxes()
        {
            if (txtStopHotkey != null)
            {
                txtStopHotkey.Text = FormatHotkey(_stopHotkeyModifiers, _stopHotkeyKey);
            }

            if (txtPauseHotkey != null)
            {
                txtPauseHotkey.Text = FormatHotkey(_pauseHotkeyModifiers, _pauseHotkeyKey);
            }
        }

        private static string FormatHotkey(uint modifiers, uint virtualKey)
        {
            var parts = new List<string>();
            if ((modifiers & MOD_CONTROL) != 0)
            {
                parts.Add("Ctrl");
            }
            if ((modifiers & MOD_SHIFT) != 0)
            {
                parts.Add("Shift");
            }
            if ((modifiers & MOD_ALT) != 0)
            {
                parts.Add("Alt");
            }
            if ((modifiers & MOD_WIN) != 0)
            {
                parts.Add("Win");
            }

            var key = KeyInterop.KeyFromVirtualKey((int)virtualKey);
            parts.Add(key == Key.None ? virtualKey.ToString() : key.ToString());
            return string.Join("+", parts);
        }

        // --- 機能1: キーを押して自動認識 ---
        private void TxtKeyDisplay_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true; // テキストボックスへの文字入力を防ぐ

            Key key = (e.Key == Key.System ? e.SystemKey : e.Key);

            // Escキーならクリア
            if (key == Key.Escape)
            {
                _targetKey = VirtualKeyCode.NONAME;
                txtKeyDisplay.Text = "（ここにフォーカスしてキーを押す）";
                SaveLastSettings();
                return;
            }

            // WPFのキーコードをInputSimulator用に変換
            int virtualKey = KeyInterop.VirtualKeyFromKey(key);
            _targetKey = (VirtualKeyCode)virtualKey;

            txtKeyDisplay.Text = $"{key} (コード: {_targetKey})";
            SaveLastSettings();
        }

        // --- 機能2: クリック位置の事前取得 (3秒タイマー) ---
        private async void BtnPickPos_Click(object sender, RoutedEventArgs e)
        {
            BtnPickPos.IsEnabled = false;
            txtPickStatus.Text = "3秒後に位置を取得します...マウスを合わせてください";

            // 3秒待機 (画面を固まらせない)
            await Task.Delay(3000);

            // 現在のマウス位置を取得 (マルチモニタ対応のグローバル座標)
            if (GetCursorPos(out POINT p))
            {
                _targetPoint = new System.Windows.Point(p.X, p.Y);
                txtX.Text = p.X.ToString();
                txtY.Text = p.Y.ToString();
                CaptureTargetWindow(p);
                await RecognizeBrowserTargetAsync(true);
                if (_browserClickTarget is null)
                {
                    txtPickStatus.Text = "位置を取得しました！";
                }
                SaveLastSettings();
            }
            else
            {
                txtPickStatus.Text = "取得に失敗しました";
                txtTargetInfo.Text = "対象: 未設定";
            }
            BtnPickPos.IsEnabled = true;
        }

        private void BtnLaunchBrowser_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var chromePath = FindChromeExecutablePath()
                    ?? throw new InvalidOperationException("Chrome が見つかりませんでした。Google Chrome をインストールしてください。");
                AppPaths.EnsureCreated();
                var profilePath = Path.Combine(AppPaths.RootPath, "dedicated-browser-profile");
                Directory.CreateDirectory(profilePath);

                var arguments = string.Join(" ", new[]
                {
                    $"--remote-debugging-port={DedicatedBrowserDebugPort}",
                    "--remote-allow-origins=*",
                    $"--user-data-dir={QuoteCommandLineArgument(profilePath)}",
                    "--no-first-run",
                    "--new-window",
                    QuoteCommandLineArgument(GetDedicatedBrowserStartUrl())
                });

                _dedicatedBrowserProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = chromePath,
                    Arguments = arguments,
                    UseShellExecute = false
                });
                txtPickStatus.Text = "専用ブラウザを起動しました。対象ページを開いてから位置を取得してください";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"専用ブラウザを起動できませんでした: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void BtnTestNonIntrusiveClick_Click(object sender, RoutedEventArgs e)
        {
            if (_timer.IsEnabled || _startDelayCts != null)
            {
                MessageBox.Show("実行中はテスト送信できません。", "確認", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(txtX.Text))
            {
                MessageBox.Show("先にクリック位置を取得してください。", "確認", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                if (!PrepareTargetWindowForClick())
                {
                    throw new InvalidOperationException("クリック位置の対象ウィンドウを取得できませんでした。");
                }

                await ClickTargetWindowAsync();
                txtPickStatus.Text = "非干渉クリックのテスト送信に成功しました";
                SaveLastSettings();
            }
            catch (Exception ex)
            {
                if (BrowserDirectClickService.IsChromiumProcess(_targetProcessName))
                {
                    MessageBox.Show(
                        $"非干渉クリックのテスト送信に失敗しました: {ex.Message}",
                        "テスト失敗",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    $"非干渉クリックのテスト送信に失敗しました: {ex.Message}\n\n物理クリック方式へ切り替えますか？",
                    "テスト失敗",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    cmbClickMethod.SelectedIndex = 1;
                    SaveLastSettings();
                }
            }
        }

        // --- スタートボタン押下 ---
        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            // バリデーション（入力チェック）
            if (cmbAction.SelectedIndex == 0)
            {
                if (!KeyboardInputService.TryParseSequence(txtKeySequence.Text, out var keySequence, out var keySequenceError))
                {
                    MessageBox.Show(keySequenceError, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (keySequence.Count == 0 && _targetKey == VirtualKeyCode.NONAME)
                {
                    MessageBox.Show("連打するキーまたはキーシーケンスが設定されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            if (cmbAction.SelectedIndex == 1 && string.IsNullOrEmpty(txtX.Text))
            {
                MessageBox.Show("クリックする位置が設定されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(txtInterval.Text, out int ms) || ms <= 0)
            {
                MessageBox.Show("間隔は正の整数(ms)で入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(txtStartDelay.Text, out int startDelaySeconds) || startDelaySeconds < 0)
            {
                MessageBox.Show("開始待機は0以上の整数(秒)で入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(txtHoldDuration.Text, out int holdDurationMs) || holdDurationMs < 1 || holdDurationMs > 60000)
            {
                MessageBox.Show("長押し時間は1から60000msの整数で入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(txtKeyHoldDuration.Text, out int keyHoldDurationMs) || keyHoldDurationMs < 1 || keyHoldDurationMs > 60000)
            {
                MessageBox.Show("キー長押し時間は1から60000msの整数で入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 継続時間のパース
            int.TryParse(txtDuration.Text, out _durationSeconds);

            if (cmbAction.SelectedIndex == 1 && cmbClickMethod.SelectedIndex == 0)
            {
                if (!PrepareTargetWindowForClick())
                {
                    MessageBox.Show("クリック位置の対象ウィンドウを取得できませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // タイマー設定
            _timer.Interval = TimeSpan.FromMilliseconds(ms);
            SaveLastSettings();

            // UI制御
            BtnStart.IsEnabled = false;
            BtnPauseResume.IsEnabled = false;
            BtnPauseResume.Content = "一時停止";
            BtnStop.IsEnabled = true;
            _isPaused = false;

            _startDelayCts?.Dispose();
            _startDelayCts = new CancellationTokenSource();
            try
            {
                await RunStartDelayAsync(startDelaySeconds, _startDelayCts.Token);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            finally
            {
                _startDelayCts?.Dispose();
                _startDelayCts = null;
            }

            _startTime = DateTime.Now;
            _timer.Start();
            _statusTimer.Start();
            BtnPauseResume.IsEnabled = true;
            UpdateRunStatus();
        }

        // --- ストップボタン押下 ---
        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            StopTimer();
        }

        private void BtnPauseResume_Click(object sender, RoutedEventArgs e)
        {
            TogglePauseResume();
        }

        private async Task RunStartDelayAsync(int startDelaySeconds, CancellationToken cancellationToken)
        {
            if (startDelaySeconds <= 0)
            {
                txtRunStatus.Text = "開始します";
                return;
            }

            for (var remaining = startDelaySeconds; remaining > 0; remaining--)
            {
                txtRunStatus.Text = $"開始待機中: {remaining}秒";
                await Task.Delay(1000, cancellationToken);
            }

            txtRunStatus.Text = "開始します";
        }

        // --- タイマー処理（ここが連打の本体） ---
        private async void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (cmbAction.SelectedIndex == 0)
                {
                    if (_isKeyInProgress)
                    {
                        return;
                    }

                    await SendTargetKeyAsync();
                }
                else
                {
                    if (_isClickInProgress)
                    {
                        return;
                    }

                    if (cmbClickMethod.SelectedIndex == 0)
                    {
                        await ClickTargetWindowWithFallbackAsync();
                    }
                    else
                    {
                        await PhysicalClickTargetPointAsync();
                    }
                }

                // 時間切れチェック
                if (_durationSeconds > 0 && (DateTime.Now - _startTime).TotalSeconds >= _durationSeconds)
                {
                    StopTimer();
                    MessageBox.Show("指定時間が経過しました。", "終了");
                }
                else
                {
                    UpdateRunStatus();
                }
            }
            catch (Exception ex)
            {
                StopTimer();
                MessageBox.Show($"エラーが発生したため停止しました: {ex.Message}", "エラー");
            }
        }

        // --- 停止処理（共通） ---
        private void StopTimer()
        {
            _startDelayCts?.Cancel();
            _timer.Stop();
            _statusTimer.Stop();
            _isPaused = false;
            BtnStart.IsEnabled = true;
            BtnPauseResume.IsEnabled = false;
            BtnPauseResume.Content = "一時停止";
            BtnStop.IsEnabled = false;
            txtRunStatus.Text = "待機中";
            SaveLastSettings();
        }

        private void TogglePauseResume()
        {
            if (_startDelayCts != null)
            {
                return;
            }

            if (_isPaused)
            {
                ResumeTimer();
                return;
            }

            if (_timer.IsEnabled)
            {
                PauseTimer();
            }
        }

        private void PauseTimer()
        {
            _timer.Stop();
            _statusTimer.Stop();
            _pausedAt = DateTime.Now;
            _isPaused = true;
            BtnPauseResume.Content = "再開";
            txtRunStatus.Text = "一時停止中";
        }

        private void ResumeTimer()
        {
            var pausedDuration = DateTime.Now - _pausedAt;
            _startTime = _startTime.Add(pausedDuration);
            _isPaused = false;
            BtnPauseResume.Content = "一時停止";
            _timer.Start();
            _statusTimer.Start();
            UpdateRunStatus();
        }

        private async Task SendTargetKeyAsync()
        {
            _isKeyInProgress = true;
            var modifiers = GetSelectedKeyModifiers();
            try
            {
                KeyboardInputService.TryParseSequence(txtKeySequence.Text, out var keySequence, out _);
                await _keyboardInputService.SendAsync(
                    _targetKey,
                    keySequence,
                    modifiers,
                    cmbKeyAction.SelectedIndex == 1,
                    GetKeyHoldDurationMs());
            }
            finally
            {
                _isKeyInProgress = false;
            }
        }

        private List<VirtualKeyCode> GetSelectedKeyModifiers()
        {
            var modifiers = new List<VirtualKeyCode>();
            if (chkKeyCtrl.IsChecked == true)
            {
                modifiers.Add(VirtualKeyCode.CONTROL);
            }
            if (chkKeyShift.IsChecked == true)
            {
                modifiers.Add(VirtualKeyCode.SHIFT);
            }
            if (chkKeyAlt.IsChecked == true)
            {
                modifiers.Add(VirtualKeyCode.MENU);
            }
            if (chkKeyWin.IsChecked == true)
            {
                modifiers.Add(VirtualKeyCode.LWIN);
            }

            return modifiers;
        }

        private int GetKeyHoldDurationMs()
        {
            return int.TryParse(txtKeyHoldDuration.Text, out var ms)
                ? Math.Clamp(ms, 1, 60000)
                : 500;
        }

        private void StatusTimer_Tick(object sender, EventArgs e)
        {
            if (_timer.IsEnabled)
            {
                UpdateRunStatus();
            }
        }

        private void UpdateRunStatus()
        {
            var elapsed = DateTime.Now - _startTime;
            var action = cmbAction.SelectedIndex == 0 ? "キー連打" : "クリック連打";

            if (_durationSeconds <= 0)
            {
                txtRunStatus.Text = $"実行中: {action} / 経過 {Math.Floor(elapsed.TotalSeconds)}秒 / 無制限";
                return;
            }

            var remaining = Math.Max(0, _durationSeconds - (int)Math.Floor(elapsed.TotalSeconds));
            txtRunStatus.Text = $"実行中: {action} / 残り {remaining}秒";
        }

        private void SaveLastSettings()
        {
            AppSettingsManager.Save(CreateCurrentSettings());
        }

        private AppSettings CreateCurrentSettings()
        {
            return new AppSettings
            {
                ModeIndex = cmbAction?.SelectedIndex ?? 0,
                ClickMethodIndex = cmbClickMethod?.SelectedIndex ?? 0,
                ClickActionIndex = cmbClickAction?.SelectedIndex ?? 0,
                HoldDurationText = txtHoldDuration?.Text ?? "500",
                AutoFallbackToPhysicalClick = chkAutoFallbackToPhysical?.IsChecked == true,
                DedicatedBrowserUrl = txtDedicatedBrowserUrl?.Text ?? string.Empty,
                TargetKey = _targetKey,
                KeyActionIndex = cmbKeyAction?.SelectedIndex ?? 0,
                KeyHoldDurationText = txtKeyHoldDuration?.Text ?? "500",
                KeySequenceText = txtKeySequence?.Text ?? string.Empty,
                KeyModifierCtrl = chkKeyCtrl?.IsChecked == true,
                KeyModifierShift = chkKeyShift?.IsChecked == true,
                KeyModifierAlt = chkKeyAlt?.IsChecked == true,
                KeyModifierWin = chkKeyWin?.IsChecked == true,
                TargetPointX = _targetPoint.X,
                TargetPointY = _targetPoint.Y,
                TargetClientPointX = _targetClientPoint.X,
                TargetClientPointY = _targetClientPoint.Y,
                HasTargetClientPoint = _hasTargetClientPoint,
                TargetWindowTitle = _targetWindowTitle,
                TargetWindowClass = _targetWindowClass,
                TargetProcessId = _targetProcessId,
                TargetProcessName = _targetProcessName,
                TargetWindowLeft = _targetWindowRect.Left,
                TargetWindowTop = _targetWindowRect.Top,
                TargetWindowRight = _targetWindowRect.Right,
                TargetWindowBottom = _targetWindowRect.Bottom,
                HasTargetWindowRect = _hasTargetWindowRect,
                IntervalText = txtInterval?.Text ?? "10000",
                DurationText = txtDuration?.Text ?? "0",
                StartDelayText = txtStartDelay?.Text ?? "0",
                StopHotkeyModifiers = _stopHotkeyModifiers,
                StopHotkeyKey = _stopHotkeyKey,
                PauseHotkeyModifiers = _pauseHotkeyModifiers,
                PauseHotkeyKey = _pauseHotkeyKey
            };
        }

        private void ApplySettings(AppSettings settings)
        {
            _isApplyingSettings = true;
            try
            {
                cmbAction.SelectedIndex = settings.ModeIndex == 1 ? 1 : 0;
                cmbClickMethod.SelectedIndex = settings.ClickMethodIndex == 1 ? 1 : 0;
                cmbClickAction.SelectedIndex = Math.Clamp(settings.ClickActionIndex, 0, 3);
                txtHoldDuration.Text = string.IsNullOrWhiteSpace(settings.HoldDurationText) ? "500" : settings.HoldDurationText;
                chkAutoFallbackToPhysical.IsChecked = settings.AutoFallbackToPhysicalClick;
                txtDedicatedBrowserUrl.Text = settings.DedicatedBrowserUrl ?? string.Empty;
                _targetKey = settings.TargetKey;
                cmbKeyAction.SelectedIndex = settings.KeyActionIndex == 1 ? 1 : 0;
                txtKeyHoldDuration.Text = string.IsNullOrWhiteSpace(settings.KeyHoldDurationText) ? "500" : settings.KeyHoldDurationText;
                txtKeySequence.Text = settings.KeySequenceText ?? string.Empty;
                chkKeyCtrl.IsChecked = settings.KeyModifierCtrl;
                chkKeyShift.IsChecked = settings.KeyModifierShift;
                chkKeyAlt.IsChecked = settings.KeyModifierAlt;
                chkKeyWin.IsChecked = settings.KeyModifierWin;
                _targetPoint = new System.Windows.Point(settings.TargetPointX, settings.TargetPointY);
                _targetClientPoint = new System.Windows.Point(settings.TargetClientPointX, settings.TargetClientPointY);
                _hasTargetClientPoint = settings.HasTargetClientPoint;
                _targetWindowTitle = settings.TargetWindowTitle ?? string.Empty;
                _targetWindowClass = settings.TargetWindowClass ?? string.Empty;
                _targetProcessId = settings.TargetProcessId;
                _targetProcessName = settings.TargetProcessName ?? string.Empty;
                RefreshTargetElevationStatus();
                _targetWindowRect = new RECT
                {
                    Left = settings.TargetWindowLeft,
                    Top = settings.TargetWindowTop,
                    Right = settings.TargetWindowRight,
                    Bottom = settings.TargetWindowBottom
                };
                _hasTargetWindowRect = settings.HasTargetWindowRect;

                if (_targetKey != VirtualKeyCode.NONAME)
                {
                    var key = KeyInterop.KeyFromVirtualKey((int)_targetKey);
                    txtKeyDisplay.Text = $"{key} (コード: {_targetKey})";
                }

                if (_targetPoint.X != 0 || _targetPoint.Y != 0)
                {
                    txtX.Text = ((int)_targetPoint.X).ToString();
                    txtY.Text = ((int)_targetPoint.Y).ToString();
                }

                txtInterval.Text = string.IsNullOrWhiteSpace(settings.IntervalText) ? "10000" : settings.IntervalText;
                txtDuration.Text = string.IsNullOrWhiteSpace(settings.DurationText) ? "0" : settings.DurationText;
                txtStartDelay.Text = string.IsNullOrWhiteSpace(settings.StartDelayText) ? "0" : settings.StartDelayText;
                _stopHotkeyModifiers = settings.StopHotkeyModifiers == 0 ? HotkeyDefaults.StopModifiers : settings.StopHotkeyModifiers;
                _stopHotkeyKey = settings.StopHotkeyKey == 0 ? HotkeyDefaults.StopKey : settings.StopHotkeyKey;
                _pauseHotkeyModifiers = settings.PauseHotkeyModifiers == 0 ? HotkeyDefaults.PauseModifiers : settings.PauseHotkeyModifiers;
                _pauseHotkeyKey = settings.PauseHotkeyKey == 0 ? HotkeyDefaults.PauseKey : settings.PauseHotkeyKey;
                UpdateHotkeyTextBoxes();
                UpdateTargetInfoText();
            }
            finally
            {
                _isApplyingSettings = false;
            }
        }

        private void CaptureTargetWindow(POINT screenPoint)
        {
            _browserClickTarget = null;
            _targetWindowHandle = WindowFromPoint(screenPoint);
            RefreshTargetWindowSnapshot();
            RefreshTargetClientPoint();

            UpdateTargetInfoText();
        }

        private bool PrepareTargetWindowForClick()
        {
            if (_targetWindowHandle != IntPtr.Zero && IsWindow(_targetWindowHandle))
            {
                RefreshTargetWindowSnapshot();
                RefreshTargetClientPoint();
                UpdateTargetInfoText();
                return _hasTargetClientPoint;
            }

            if (TryRecaptureTargetWindowFromSavedPoint())
            {
                UpdateTargetInfoText();
                return _hasTargetClientPoint;
            }

            if (HasTargetWindowIdentity())
            {
                var matchedWindow = FindBestTargetWindow();
                if (matchedWindow != IntPtr.Zero)
                {
                    _targetWindowHandle = matchedWindow;
                    RefreshTargetWindowSnapshot();
                    RefreshTargetClientPoint();
                    UpdateTargetInfoText();
                    return _hasTargetClientPoint;
                }
            }

            var screenPoint = new POINT
            {
                X = (int)_targetPoint.X,
                Y = (int)_targetPoint.Y
            };
            CaptureTargetWindow(screenPoint);
            return _targetWindowHandle != IntPtr.Zero && _hasTargetClientPoint;
        }

        private bool TryRecaptureTargetWindowFromSavedPoint()
        {
            if (_targetPoint.X == 0 && _targetPoint.Y == 0)
            {
                return false;
            }

            var screenPoint = new POINT
            {
                X = (int)_targetPoint.X,
                Y = (int)_targetPoint.Y
            };
            var candidateWindow = WindowFromPoint(screenPoint);
            if (candidateWindow == IntPtr.Zero
                || !IsWindow(candidateWindow)
                || GetWindowProcessId(candidateWindow) == Process.GetCurrentProcess().Id)
            {
                return false;
            }

            if (HasTargetWindowIdentity() && ScoreTargetWindow(candidateWindow) < 50)
            {
                return false;
            }

            _targetWindowHandle = candidateWindow;
            RefreshTargetWindowSnapshot();
            RefreshTargetClientPoint();
            return _hasTargetClientPoint;
        }

        private async Task RecognizeBrowserTargetAsync(bool updateStatus)
        {
            _browserClickTarget = null;
            if (!BrowserDirectClickService.IsChromiumProcess(_targetProcessName))
            {
                UpdateTargetInfoText();
                return;
            }

            var target = await _browserDirectClickService.TryResolveTargetAsync(
                _targetWindowTitle,
                _targetProcessName,
                DedicatedBrowserDebugPort,
                CancellationToken.None);
            if (target is null)
            {
                if (updateStatus)
                {
                    txtPickStatus.Text = "Chrome系ウィンドウですが、専用ブラウザとして認識できませんでした";
                }

                UpdateTargetInfoText();
                return;
            }

            _browserClickTarget = target;
            if (updateStatus)
            {
                txtPickStatus.Text = $"専用ブラウザを認識しました: {target.DisplayName}";
            }

            UpdateTargetInfoText();
        }

        private void RefreshTargetWindowSnapshot()
        {
            if (_targetWindowHandle == IntPtr.Zero)
            {
                _targetWindowTitle = string.Empty;
                _targetWindowClass = string.Empty;
                _targetProcessId = 0;
                _targetProcessName = string.Empty;
                _hasTargetWindowRect = false;
                RefreshTargetElevationStatus();
                return;
            }

            _targetWindowTitle = GetWindowTitle(_targetWindowHandle);
            _targetWindowClass = GetWindowClass(_targetWindowHandle);
            _targetProcessId = GetWindowProcessId(_targetWindowHandle);
            _targetProcessName = GetProcessName(_targetProcessId);
            var rootWindow = GetAncestor(_targetWindowHandle, GA_ROOT);
            var rootTitle = GetWindowTitle(rootWindow);
            if (!string.IsNullOrWhiteSpace(rootTitle)
                && (string.IsNullOrWhiteSpace(_targetWindowTitle)
                    || BrowserDirectClickService.IsChromiumProcess(_targetProcessName)
                    || _targetWindowClass.Contains("Chrome", StringComparison.OrdinalIgnoreCase)))
            {
                _targetWindowTitle = rootTitle;
            }

            RefreshTargetElevationStatus();
            _hasTargetWindowRect = GetWindowRect(_targetWindowHandle, out _targetWindowRect);
        }

        private void RefreshTargetClientPoint()
        {
            var clientPoint = new POINT
            {
                X = (int)_targetPoint.X,
                Y = (int)_targetPoint.Y
            };
            _hasTargetClientPoint = _targetWindowHandle != IntPtr.Zero
                && ScreenToClient(_targetWindowHandle, ref clientPoint);
            _targetClientPoint = _hasTargetClientPoint
                ? new System.Windows.Point(clientPoint.X, clientPoint.Y)
                : new System.Windows.Point();
        }

        private async Task ClickTargetWindowAsync()
        {
            if (!PrepareTargetWindowForClick())
            {
                throw new InvalidOperationException("クリック対象ウィンドウを再取得できませんでした。");
            }

            if (!_hasTargetClientPoint)
            {
                throw new InvalidOperationException("クリック対象のウィンドウ内座標が設定されていません。");
            }

            var clientPoint = new POINT
            {
                X = (int)_targetClientPoint.X,
                Y = (int)_targetClientPoint.Y
            };
            var clickAction = GetSelectedClickAction();
            var holdDurationMs = GetHoldDurationMs();
            _isClickInProgress = true;
            try
            {
                if (BrowserDirectClickService.IsChromiumProcess(_targetProcessName))
                {
                    if (_browserClickTarget is null || !_browserClickTarget.IsRecognized)
                    {
                        await RecognizeBrowserTargetAsync(false);
                    }

                    if (_browserClickTarget is null || !_browserClickTarget.IsRecognized)
                    {
                        throw new InvalidOperationException("Chrome系の対象は専用ブラウザとして認識されていません。専用ブラウザを起動し、その中で対象ページを開いてから位置を取得してください。");
                    }

                    if (await _browserDirectClickService.TryClickAsync(
                        _browserClickTarget,
                        clientPoint,
                        clickAction,
                        holdDurationMs,
                        CancellationToken.None))
                    {
                        txtPickStatus.Text = $"専用ブラウザへ直接クリックを送信しました: {_browserClickTarget.DisplayName}";
                        UpdateTargetInfoText();
                        return;
                    }

                    throw new InvalidOperationException("認識済みの専用ブラウザへ直接クリックを送信できませんでした。対象タブを開き直してから位置を取得してください。");
                }

                await _mouseClickService.SendNonIntrusiveAsync(
                    _targetWindowHandle,
                    clientPoint,
                    clickAction,
                    holdDurationMs);
            }
            finally
            {
                _isClickInProgress = false;
            }
        }

        private async Task ClickTargetWindowWithFallbackAsync()
        {
            try
            {
                await ClickTargetWindowAsync();
            }
            catch
            {
                if (BrowserDirectClickService.IsChromiumProcess(_targetProcessName))
                {
                    throw;
                }

                if (chkAutoFallbackToPhysical.IsChecked != true)
                {
                    throw;
                }

                txtPickStatus.Text = "非干渉クリックに失敗したため、許可設定に従って物理クリックを実行しました";
                await PhysicalClickTargetPointAsync();
            }
        }

        private async Task PhysicalClickTargetPointAsync()
        {
            _isClickInProgress = true;
            try
            {
                var targetPoint = new POINT
                {
                    X = (int)_targetPoint.X,
                    Y = (int)_targetPoint.Y
                };
                await _mouseClickService.SendPhysicalAsync(
                    targetPoint,
                    GetSelectedClickAction(),
                    GetHoldDurationMs());
            }
            finally
            {
                _isClickInProgress = false;
            }
        }

        private MouseClickAction GetSelectedClickAction()
        {
            return cmbClickAction.SelectedIndex switch
            {
                1 => MouseClickAction.Right,
                2 => MouseClickAction.DoubleLeft,
                3 => MouseClickAction.HoldLeft,
                _ => MouseClickAction.Left
            };
        }

        private int GetHoldDurationMs()
        {
            return int.TryParse(txtHoldDuration.Text, out var ms)
                ? Math.Clamp(ms, 1, 60000)
                : 500;
        }

        private void UpdateTargetInfoText()
        {
            if (txtTargetInfo == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(_targetWindowTitle) && string.IsNullOrWhiteSpace(_targetWindowClass))
            {
                txtTargetInfo.Text = "対象: 未設定";
                return;
            }

            var title = string.IsNullOrWhiteSpace(_targetWindowTitle) ? "タイトルなし" : _targetWindowTitle;
            var className = string.IsNullOrWhiteSpace(_targetWindowClass) ? "クラス不明" : _targetWindowClass;
            var process = string.IsNullOrWhiteSpace(_targetProcessName)
                ? "プロセス不明"
                : $"{_targetProcessName}({_targetProcessId})";
            var client = _hasTargetClientPoint
                ? $" / Client X={(int)_targetClientPoint.X}, Y={(int)_targetClientPoint.Y}"
                : string.Empty;
            var elevation = string.IsNullOrWhiteSpace(_targetElevationStatus)
                ? "権限不明"
                : _targetElevationStatus;
            var warning = _targetMayRequireElevation
                ? " / 入力制限の可能性"
                : string.Empty;
            var browser = string.Empty;
            if (BrowserDirectClickService.IsChromiumProcess(_targetProcessName))
            {
                browser = _browserClickTarget is not null && _browserClickTarget.IsRecognized
                    ? $" / 専用ブラウザ認識: {_browserClickTarget.DisplayName}"
                    : " / 専用ブラウザ未認識";
            }

            txtTargetInfo.Text = $"対象: {title} / {className} / {process} / {elevation}{warning}{client}{browser}";
        }

        private bool HasTargetWindowIdentity()
        {
            return _targetProcessId > 0
                || !string.IsNullOrWhiteSpace(_targetProcessName)
                || !string.IsNullOrWhiteSpace(_targetWindowTitle)
                || !string.IsNullOrWhiteSpace(_targetWindowClass)
                || _hasTargetWindowRect;
        }

        private IntPtr FindBestTargetWindow()
        {
            var bestWindow = IntPtr.Zero;
            var bestScore = 0;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }

                var score = ScoreTargetWindow(hWnd);
                if (score <= bestScore)
                {
                    return true;
                }

                bestScore = score;
                bestWindow = hWnd;
                return true;
            }, IntPtr.Zero);

            return bestScore >= 60 ? bestWindow : IntPtr.Zero;
        }

        private int ScoreTargetWindow(IntPtr hWnd)
        {
            var score = 0;
            var processId = GetWindowProcessId(hWnd);
            var processName = GetProcessName(processId);
            var windowTitle = GetWindowTitle(hWnd);
            var windowClass = GetWindowClass(hWnd);

            if (_targetProcessId > 0 && processId == _targetProcessId)
            {
                score += 45;
            }

            if (!string.IsNullOrWhiteSpace(_targetWindowClass)
                && string.Equals(windowClass, _targetWindowClass, StringComparison.Ordinal))
            {
                score += 35;
            }

            if (!string.IsNullOrWhiteSpace(_targetWindowTitle)
                && string.Equals(windowTitle, _targetWindowTitle, StringComparison.Ordinal))
            {
                score += 30;
            }

            if (!string.IsNullOrWhiteSpace(_targetProcessName)
                && string.Equals(processName, _targetProcessName, StringComparison.OrdinalIgnoreCase))
            {
                score += 20;
            }

            if (_hasTargetWindowRect
                && GetWindowRect(hWnd, out var rect)
                && IsSimilarRect(rect, _targetWindowRect))
            {
                score += 15;
            }

            return score;
        }

        private static bool IsSimilarRect(RECT current, RECT expected)
        {
            const int tolerance = 24;
            return Math.Abs(current.Left - expected.Left) <= tolerance
                && Math.Abs(current.Top - expected.Top) <= tolerance
                && Math.Abs(current.Right - expected.Right) <= tolerance
                && Math.Abs(current.Bottom - expected.Bottom) <= tolerance;
        }

        private void RefreshTargetElevationStatus()
        {
            _targetElevationStatus = "権限不明";
            _targetMayRequireElevation = false;

            if (_targetProcessId <= 0)
            {
                return;
            }

            if (TryGetProcessElevation(_targetProcessId, out var isElevated, out var accessDenied))
            {
                _targetElevationStatus = isElevated ? "管理者権限" : "通常権限";
                _targetMayRequireElevation = isElevated && !_isCurrentProcessElevated;
                return;
            }

            if (accessDenied)
            {
                _targetElevationStatus = "権限不明（管理者権限の可能性）";
                _targetMayRequireElevation = !_isCurrentProcessElevated;
            }
        }

        private static bool TryGetProcessElevation(int processId, out bool isElevated, out bool accessDenied)
        {
            isElevated = false;
            accessDenied = false;

            if (processId <= 0)
            {
                return false;
            }

            var processHandle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, (uint)processId);
            if (processHandle == IntPtr.Zero)
            {
                accessDenied = Marshal.GetLastWin32Error() == ERROR_ACCESS_DENIED;
                return false;
            }

            IntPtr tokenHandle = IntPtr.Zero;
            try
            {
                if (!OpenProcessToken(processHandle, TOKEN_QUERY, out tokenHandle))
                {
                    accessDenied = Marshal.GetLastWin32Error() == ERROR_ACCESS_DENIED;
                    return false;
                }

                var tokenElevation = new TOKEN_ELEVATION();
                var tokenSize = Marshal.SizeOf<TOKEN_ELEVATION>();
                if (!GetTokenInformation(tokenHandle, TokenElevation, out tokenElevation, tokenSize, out _))
                {
                    accessDenied = Marshal.GetLastWin32Error() == ERROR_ACCESS_DENIED;
                    return false;
                }

                isElevated = tokenElevation.TokenIsElevated != 0;
                return true;
            }
            finally
            {
                if (tokenHandle != IntPtr.Zero)
                {
                    CloseHandle(tokenHandle);
                }

                CloseHandle(processHandle);
            }
        }

        private static int GetWindowProcessId(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return 0;
            }

            GetWindowThreadProcessId(hWnd, out var processId);
            return (int)processId;
        }

        private static string GetProcessName(int processId)
        {
            if (processId <= 0)
            {
                return string.Empty;
            }

            try
            {
                using var process = Process.GetProcessById(processId);
                return process.ProcessName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string FindChromeExecutablePath()
        {
            var candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Google", "Chrome", "Application", "chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Google", "Chrome", "Application", "chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google", "Chrome", "Application", "chrome.exe")
            };

            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static string QuoteCommandLineArgument(string value)
        {
            return $"\"{value.Replace("\"", "\\\"", StringComparison.Ordinal)}\"";
        }

        private string GetDedicatedBrowserStartUrl()
        {
            var rawUrl = txtDedicatedBrowserUrl?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(rawUrl))
            {
                return "about:blank";
            }

            if (Uri.TryCreate(rawUrl, UriKind.Absolute, out var uri))
            {
                return uri.ToString();
            }

            return $"https://{rawUrl}";
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

        private static string GetWindowTitle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return string.Empty;
            }

            var length = GetWindowTextLength(hWnd);
            var builder = new StringBuilder(length + 1);
            GetWindowText(hWnd, builder, builder.Capacity);
            return builder.ToString();
        }

        private static string GetWindowClass(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return string.Empty;
            }

            var builder = new StringBuilder(256);
            GetClassName(hWnd, builder, builder.Capacity);
            return builder.ToString();
        }

    }
}


