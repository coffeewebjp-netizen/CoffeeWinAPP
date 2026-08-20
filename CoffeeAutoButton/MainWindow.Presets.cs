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
    public partial class MainWindow
    {
        private void LoadPresetsToUI()
        {
            cmbPresets.ItemsSource = null;
            cmbPresets.ItemsSource = _presets;
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
    }
}
