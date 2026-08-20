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
    }
}
