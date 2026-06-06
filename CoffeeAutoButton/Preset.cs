using GregsStack.InputSimulatorStandard.Native;
using System.Windows;

namespace CoffeeAutoButton
{
    public class Preset
    {
        public int SchemaVersion { get; set; } = AppSchema.CurrentVersion;
        public string Name { get; set; }
        public int ModeIndex { get; set; } // 0: Key, 1: Click
        public int ClickMethodIndex { get; set; } // 0: Non-intrusive, 1: Physical
        public int ClickActionIndex { get; set; } // 0: Left, 1: Right, 2: Double left, 3: Hold left
        public string HoldDurationText { get; set; }
        public bool AutoFallbackToPhysicalClick { get; set; }
        public string DedicatedBrowserUrl { get; set; }
        public VirtualKeyCode TargetKey { get; set; }
        public int KeyActionIndex { get; set; }
        public string KeyHoldDurationText { get; set; }
        public string KeySequenceText { get; set; }
        public bool KeyModifierCtrl { get; set; }
        public bool KeyModifierShift { get; set; }
        public bool KeyModifierAlt { get; set; }
        public bool KeyModifierWin { get; set; }
        public Point TargetPoint { get; set; }
        public Point TargetClientPoint { get; set; }
        public bool HasTargetClientPoint { get; set; }
        public string TargetWindowTitle { get; set; }
        public string TargetWindowClass { get; set; }
        public int TargetProcessId { get; set; }
        public string TargetProcessName { get; set; }
        public int TargetWindowLeft { get; set; }
        public int TargetWindowTop { get; set; }
        public int TargetWindowRight { get; set; }
        public int TargetWindowBottom { get; set; }
        public bool HasTargetWindowRect { get; set; }
        public string IntervalText { get; set; }
        public string DurationText { get; set; }
        public string StartDelayText { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}

