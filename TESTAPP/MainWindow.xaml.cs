using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using GregsStack.InputSimulatorStandard;
using GregsStack.InputSimulatorStandard.Native;
using System.Windows.Interop;

namespace AutoKeyInputApp
{
    public partial class MainWindow : Window
    {
        private readonly InputSimulator _simulator = new InputSimulator();
        private readonly DispatcherTimer _timer = new DispatcherTimer();
        private VirtualKeyCode _keyToPress;
        private DateTime _startTime;
        private int _durationSeconds;

        public MainWindow()
        {
            InitializeComponent();
            _timer.Tick += Timer_Tick;
        }

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            // ① 動作選択
            bool isKey = (cmbAction.SelectedIndex == 0);

            // ② キー入力の場合はパース
            if (isKey)
            {
                string raw = txtKey.Text.Trim();
                if (!TryGetVirtualKeyCode(raw, out _keyToPress))
                {
                    MessageBox.Show($"無効なキーです: 「{raw}」", "入力エラー",
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // ③ 間隔パース
            if (!int.TryParse(txtInterval.Text.Trim(), out int ms) || ms <= 0)
            {
                MessageBox.Show("間隔には正の整数（ミリ秒）を入力してください。", "入力エラー",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ④ 継続時間パース（0以下なら無制限）
            if (!int.TryParse(txtDuration.Text.Trim(), out _durationSeconds) || _durationSeconds < 1)
            {
                _durationSeconds = 0;
            }

            // ⑤ タイマー開始情報をセット
            _timer.Interval = TimeSpan.FromMilliseconds(ms);
            _startTime = DateTime.Now;
            _timer.Start();

            // ⑥ UI状態変更
            BtnStart.IsEnabled = false;
            BtnStop.IsEnabled = true;
            BtnClear.IsEnabled = false;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // 動作実行
            if (cmbAction.SelectedIndex == 0)
                _simulator.Keyboard.KeyPress(_keyToPress);
            else
                _simulator.Mouse.LeftButtonClick();

            // 継続時間チェック
            if (_durationSeconds > 0 &&
                (DateTime.Now - _startTime).TotalSeconds >= _durationSeconds)
            {
                _timer.Stop();
                BtnStart.IsEnabled = true;
                BtnStop.IsEnabled = false;
                BtnClear.IsEnabled = true;
                MessageBox.Show("指定時間が経過したため停止しました。", "情報",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            _timer.Stop();
            BtnStart.IsEnabled = true;
            BtnStop.IsEnabled = false;
            BtnClear.IsEnabled = true;
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            // 入力欄と選択をリセット
            cmbAction.SelectedIndex = 0;
            txtKey.Clear();
            txtInterval.Clear();
            txtDuration.Clear();
        }

        // キー文字列 → VirtualKeyCode 変換
        private bool TryGetVirtualKeyCode(string input, out VirtualKeyCode vk)
        {
            vk = default;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            // 1文字なら "VK_" を付与
            string token = (input.Length == 1 && char.IsLetter(input[0]))
                           ? "VK_" + input.ToUpper()
                           : input.ToUpper();

            if (Enum.TryParse(token, out vk))
                return true;

            if (Enum.TryParse(input, true, out Key wpfKey))
            {
                int code = KeyInterop.VirtualKeyFromKey(wpfKey);
                vk = (VirtualKeyCode)code;
                return true;
            }

            return false;
        }
    }
}
