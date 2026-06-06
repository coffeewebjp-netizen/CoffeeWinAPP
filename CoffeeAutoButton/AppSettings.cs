using GregsStack.InputSimulatorStandard.Native;

namespace CoffeeAutoButton
{
    public class AppSettings
    {
        public int SchemaVersion { get; set; } = AppSchema.CurrentVersion;
        public int ModeIndex { get; set; }
        public int ClickMethodIndex { get; set; }
        public int ClickActionIndex { get; set; }
        public string HoldDurationText { get; set; } = "500";
        public bool AutoFallbackToPhysicalClick { get; set; }
        public string DedicatedBrowserUrl { get; set; } = string.Empty;
        public VirtualKeyCode TargetKey { get; set; }
        public int KeyActionIndex { get; set; }
        public string KeyHoldDurationText { get; set; } = "500";
        public string KeySequenceText { get; set; } = string.Empty;
        public bool KeyModifierCtrl { get; set; }
        public bool KeyModifierShift { get; set; }
        public bool KeyModifierAlt { get; set; }
        public bool KeyModifierWin { get; set; }
        public double TargetPointX { get; set; }
        public double TargetPointY { get; set; }
        public double TargetClientPointX { get; set; }
        public double TargetClientPointY { get; set; }
        public bool HasTargetClientPoint { get; set; }
        public string TargetWindowTitle { get; set; } = string.Empty;
        public string TargetWindowClass { get; set; } = string.Empty;
        public int TargetProcessId { get; set; }
        public string TargetProcessName { get; set; } = string.Empty;
        public int TargetWindowLeft { get; set; }
        public int TargetWindowTop { get; set; }
        public int TargetWindowRight { get; set; }
        public int TargetWindowBottom { get; set; }
        public bool HasTargetWindowRect { get; set; }
        public string IntervalText { get; set; } = "10000";
        public string DurationText { get; set; } = "0";
        public string StartDelayText { get; set; } = "0";
        public uint StopHotkeyModifiers { get; set; } = HotkeyDefaults.StopModifiers;
        public uint StopHotkeyKey { get; set; } = HotkeyDefaults.StopKey;
        public uint PauseHotkeyModifiers { get; set; } = HotkeyDefaults.PauseModifiers;
        public uint PauseHotkeyKey { get; set; } = HotkeyDefaults.PauseKey;
    }
}
