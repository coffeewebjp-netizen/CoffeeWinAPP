using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Input;
using GregsStack.InputSimulatorStandard;
using GregsStack.InputSimulatorStandard.Native;

namespace CoffeeAutoButton
{
    internal sealed class KeyboardInputService
    {
        private readonly InputSimulator _simulator = new InputSimulator();

        internal async Task SendAsync(
            VirtualKeyCode targetKey,
            IReadOnlyList<VirtualKeyCode> sequence,
            IReadOnlyList<VirtualKeyCode> modifiers,
            bool isHold,
            int holdDurationMs)
        {
            foreach (var modifier in modifiers)
            {
                _simulator.Keyboard.KeyDown(modifier);
            }

            try
            {
                if (sequence.Count > 0)
                {
                    foreach (var key in sequence)
                    {
                        _simulator.Keyboard.KeyPress(key);
                        await Task.Delay(30);
                    }
                }
                else if (isHold)
                {
                    _simulator.Keyboard.KeyDown(targetKey);
                    await Task.Delay(holdDurationMs);
                    _simulator.Keyboard.KeyUp(targetKey);
                }
                else
                {
                    _simulator.Keyboard.KeyPress(targetKey);
                }
            }
            finally
            {
                for (var index = modifiers.Count - 1; index >= 0; index--)
                {
                    _simulator.Keyboard.KeyUp(modifiers[index]);
                }
            }
        }

        internal static bool TryParseSequence(string text, out List<VirtualKeyCode> sequence, out string error)
        {
            sequence = new List<VirtualKeyCode>();
            error = string.Empty;

            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            var tokens = text.Split(
                new[] { ',', ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var token in tokens)
            {
                if (!TryParseKeyToken(token, out var virtualKey))
                {
                    error = $"キーシーケンスの「{token}」を認識できません。";
                    return false;
                }

                sequence.Add(virtualKey);
            }

            return true;
        }

        private static bool TryParseKeyToken(string token, out VirtualKeyCode virtualKey)
        {
            virtualKey = VirtualKeyCode.NONAME;
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            var normalized = token.Trim();
            Key key;
            switch (normalized.ToUpperInvariant())
            {
                case "ENTER":
                case "RETURN":
                    key = Key.Return;
                    break;
                case "ESC":
                    key = Key.Escape;
                    break;
                case "SPACE":
                    key = Key.Space;
                    break;
                case "TAB":
                    key = Key.Tab;
                    break;
                case "BACKSPACE":
                case "BKSP":
                    key = Key.Back;
                    break;
                case "DEL":
                case "DELETE":
                    key = Key.Delete;
                    break;
                case "INS":
                case "INSERT":
                    key = Key.Insert;
                    break;
                case "UP":
                    key = Key.Up;
                    break;
                case "DOWN":
                    key = Key.Down;
                    break;
                case "LEFT":
                    key = Key.Left;
                    break;
                case "RIGHT":
                    key = Key.Right;
                    break;
                case "PGUP":
                case "PAGEUP":
                    key = Key.PageUp;
                    break;
                case "PGDN":
                case "PAGEDOWN":
                    key = Key.PageDown;
                    break;
                default:
                    if (normalized.Length == 1 && char.IsLetter(normalized[0]))
                    {
                        key = (Key)((int)Key.A + (char.ToUpperInvariant(normalized[0]) - 'A'));
                    }
                    else if (normalized.Length == 1 && char.IsDigit(normalized[0]))
                    {
                        key = (Key)((int)Key.D0 + (normalized[0] - '0'));
                    }
                    else if (!Enum.TryParse(normalized, ignoreCase: true, out key))
                    {
                        return false;
                    }
                    break;
            }

            var keyCode = KeyInterop.VirtualKeyFromKey(key);
            if (keyCode == 0)
            {
                return false;
            }

            virtualKey = (VirtualKeyCode)keyCode;
            return true;
        }
    }
}
