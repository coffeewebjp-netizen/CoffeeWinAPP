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

            var screenPoint = new POINT
            {
                X = (int)_targetPoint.X,
                Y = (int)_targetPoint.Y
            };
            var targetWindowRect = _hasTargetWindowRect
                ? _targetWindowRect
                : (RECT?)null;
            var target = await _browserDirectClickService.TryResolveTargetAsync(
                _targetWindowTitle,
                _targetProcessName,
                DedicatedBrowserDebugPort,
                screenPoint,
                targetWindowRect,
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
