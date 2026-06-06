using System;
using System.Threading.Tasks;
using static CoffeeAutoButton.NativeMethods;

namespace CoffeeAutoButton
{
    internal enum MouseClickAction
    {
        Left = 0,
        Right = 1,
        DoubleLeft = 2,
        HoldLeft = 3
    }

    internal sealed class MouseClickService
    {
        internal async Task SendNonIntrusiveAsync(
            IntPtr targetWindowHandle,
            POINT clientPoint,
            MouseClickAction action,
            int holdDurationMs)
        {
            var lParam = MakeLParam(clientPoint.X, clientPoint.Y);
            if (!PostMessage(targetWindowHandle, WM_MOUSEMOVE, IntPtr.Zero, lParam))
            {
                throw new InvalidOperationException("クリックメッセージの送信に失敗しました。");
            }

            switch (action)
            {
                case MouseClickAction.Right:
                    SendNonIntrusiveButton(targetWindowHandle, WM_RBUTTONDOWN, WM_RBUTTONUP, MK_RBUTTON, lParam);
                    break;
                case MouseClickAction.DoubleLeft:
                    SendNonIntrusiveButton(targetWindowHandle, WM_LBUTTONDOWN, WM_LBUTTONUP, MK_LBUTTON, lParam);
                    await Task.Delay(80);
                    if (!PostMessage(targetWindowHandle, WM_LBUTTONDBLCLK, new IntPtr(MK_LBUTTON), lParam)
                        || !PostMessage(targetWindowHandle, WM_LBUTTONUP, IntPtr.Zero, lParam))
                    {
                        throw new InvalidOperationException("ダブルクリックメッセージの送信に失敗しました。");
                    }
                    break;
                case MouseClickAction.HoldLeft:
                    if (!PostMessage(targetWindowHandle, WM_LBUTTONDOWN, new IntPtr(MK_LBUTTON), lParam))
                    {
                        throw new InvalidOperationException("長押し開始メッセージの送信に失敗しました。");
                    }
                    await Task.Delay(holdDurationMs);
                    if (!PostMessage(targetWindowHandle, WM_LBUTTONUP, IntPtr.Zero, lParam))
                    {
                        throw new InvalidOperationException("長押し終了メッセージの送信に失敗しました。");
                    }
                    break;
                default:
                    SendNonIntrusiveButton(targetWindowHandle, WM_LBUTTONDOWN, WM_LBUTTONUP, MK_LBUTTON, lParam);
                    break;
            }
        }

        internal async Task SendPhysicalAsync(POINT targetPoint, MouseClickAction action, int holdDurationMs)
        {
            if (!GetCursorPos(out var originalPosition))
            {
                throw new InvalidOperationException("現在のマウス位置を取得できませんでした。");
            }

            try
            {
                SetCursorPos(targetPoint.X, targetPoint.Y);
                await SendPhysicalClickAsync(action, holdDurationMs);
            }
            finally
            {
                SetCursorPos(originalPosition.X, originalPosition.Y);
            }
        }

        private static async Task SendPhysicalClickAsync(MouseClickAction action, int holdDurationMs)
        {
            switch (action)
            {
                case MouseClickAction.Right:
                    SendPhysicalButton(MouseEventRightDown, MouseEventRightUp);
                    break;
                case MouseClickAction.DoubleLeft:
                    SendPhysicalButton(MouseEventLeftDown, MouseEventLeftUp);
                    await Task.Delay(80);
                    SendPhysicalButton(MouseEventLeftDown, MouseEventLeftUp);
                    break;
                case MouseClickAction.HoldLeft:
                    mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
                    await Task.Delay(holdDurationMs);
                    mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
                    break;
                default:
                    SendPhysicalButton(MouseEventLeftDown, MouseEventLeftUp);
                    break;
            }
        }

        private static void SendNonIntrusiveButton(
            IntPtr targetWindowHandle,
            uint downMessage,
            uint upMessage,
            int buttonState,
            IntPtr lParam)
        {
            if (!PostMessage(targetWindowHandle, downMessage, new IntPtr(buttonState), lParam)
                || !PostMessage(targetWindowHandle, upMessage, IntPtr.Zero, lParam))
            {
                throw new InvalidOperationException("クリックメッセージの送信に失敗しました。");
            }
        }

        private static void SendPhysicalButton(uint downFlag, uint upFlag)
        {
            mouse_event(downFlag, 0, 0, 0, UIntPtr.Zero);
            mouse_event(upFlag, 0, 0, 0, UIntPtr.Zero);
        }

        private static IntPtr MakeLParam(int lowWord, int highWord)
        {
            return new IntPtr(((ushort)highWord << 16) | ((ushort)lowWord));
        }
    }
}
