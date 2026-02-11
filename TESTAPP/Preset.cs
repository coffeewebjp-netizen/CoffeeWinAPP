using GregsStack.InputSimulatorStandard.Native;
using System.Windows;

namespace AutoKeyInputApp
{
    public class Preset
    {
        public string Name { get; set; }
        public int ModeIndex { get; set; } // 0: Key, 1: Click
        public VirtualKeyCode TargetKey { get; set; }
        public Point TargetPoint { get; set; }
        public string IntervalText { get; set; }
        public string DurationText { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
