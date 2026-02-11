using System;
using System.Runtime.InteropServices; // Windows API利用に必要
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using GregsStack.InputSimulatorStandard;
using GregsStack.InputSimulatorStandard.Native;

namespace AutoKeyInputApp
{
    public partial class MainWindow : Window
    {
        private readonly InputSimulator _simulator = new InputSimulator();
        private readonly DispatcherTimer _timer = new DispatcherTimer();

        // 実行用変数
        private VirtualKeyCode _targetKey = VirtualKeyCode.NONAME;
        private System.Windows.Point _targetPoint;
        private DateTime _startTime;
        private int _durationSeconds;

        // --- Windows API (外部DLL) の読み込み ---
        // マウスの座標を取得・設定するために user32.dll の機能を使います

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetCursorPos(int x, int y);

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT { public int X; public int Y; }
        // ---------------------------------------

        public MainWindow()
        {
            InitializeComponent();
            _timer.Tick += Timer_Tick;
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
                return;
            }

            // WPFのキーコードをInputSimulator用に変換
            int virtualKey = KeyInterop.VirtualKeyFromKey(key);
            _targetKey = (VirtualKeyCode)virtualKey;

            txtKeyDisplay.Text = $"{key} (コード: {_targetKey})";
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
                txtPickStatus.Text = "位置を取得しました！";
            }
            else
            {
                txtPickStatus.Text = "取得に失敗しました";
            }
            BtnPickPos.IsEnabled = true;
        }

        // --- スタートボタン押下 ---
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            // バリデーション（入力チェック）
            if (cmbAction.SelectedIndex == 0 && _targetKey == VirtualKeyCode.NONAME)
            {
                MessageBox.Show("連打するキーが設定されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
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

            // 継続時間のパース
            int.TryParse(txtDuration.Text, out _durationSeconds);

            // タイマー設定
            _timer.Interval = TimeSpan.FromMilliseconds(ms);
            _startTime = DateTime.Now;
            _timer.Start();

            // UI制御
            BtnStart.IsEnabled = false;
            BtnStop.IsEnabled = true;
            this.Topmost = true; // 実行中はウィンドウを最前面固定
        }

        // --- ストップボタン押下 ---
        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            StopTimer();
        }

        // --- タイマー処理（ここが連打の本体） ---
        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (cmbAction.SelectedIndex == 0)
                {
                    // キー連打
                    _simulator.Keyboard.KeyPress(_targetKey);
                }
                else
                {
                    // クリック連打
                    // 1. 指定座標へ瞬間移動 (SetCursorPosを使用することでマルチモニタ対応)
                    SetCursorPos((int)_targetPoint.X, (int)_targetPoint.Y);

                    // 2. 左クリック実行
                    _simulator.Mouse.LeftButtonClick();
                }

                // 時間切れチェック
                if (_durationSeconds > 0 && (DateTime.Now - _startTime).TotalSeconds >= _durationSeconds)
                {
                    StopTimer();
                    MessageBox.Show("指定時間が経過しました。", "終了");
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
            _timer.Stop();
            BtnStart.IsEnabled = true;
            BtnStop.IsEnabled = false;
            this.Topmost = false; // 最前面固定を解除
        }
    }
}