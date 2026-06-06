namespace CoffeeAutoButton
{
    internal static class HotkeyDefaults
    {
        internal const uint StopKey = 0x53;
        internal const uint PauseKey = 0x50;
        internal const uint StopModifiers = NativeMethods.MOD_CONTROL | NativeMethods.MOD_SHIFT;
        internal const uint PauseModifiers = NativeMethods.MOD_CONTROL | NativeMethods.MOD_SHIFT;
    }
}
