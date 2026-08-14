using Microsoft.Win32;
using System;

namespace RecallTape.OneNote
{
    /// <summary>
    /// User settings, in HKCU\Software\EWC3 Labs\RecallTape.
    ///
    /// The registry rather than a config file, deliberately: the add-in lives in Program Files,
    /// which is read-only to the user who is running it, and a settings file next to the binary
    /// would either fail to save or force the install directory to be writable. Writable install
    /// directories are the thing we just moved away from.
    ///
    /// Values are cached after first read. Nothing here is hot enough to justify a registry hit per
    /// call, and Log() is called on every operation.
    /// </summary>
    internal static class Settings
    {
        private const string Key = @"Software\EWC3 Labs\RecallTape";

        private static bool loaded;
        private static bool developerTools;
        private static bool logging = true;
        private static int logMaxKB = 1024;

        private static void Load()
        {
            if (loaded) return;
            loaded = true;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(Key))
                {
                    if (k == null) return;
                    developerTools = Convert.ToInt32(k.GetValue("DeveloperTools", 0)) != 0;
                    logging = Convert.ToInt32(k.GetValue("Logging", 1)) != 0;
                    logMaxKB = Convert.ToInt32(k.GetValue("LogMaxKB", 1024));
                }
            }
            catch { /* defaults are fine; settings must never stop the add-in loading */ }
        }

        private static void Save(string name, int value)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(Key))
                {
                    if (k != null) k.SetValue(name, value, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        /// <summary>Show Dump Page XML and Survey Notebooks. Off for people who just study.</summary>
        public static bool DeveloperTools
        {
            get { Load(); return developerTools; }
            set { Load(); developerTools = value; Save("DeveloperTools", value ? 1 : 0); }
        }

        /// <summary>Write recalltape.log at all.</summary>
        public static bool Logging
        {
            get { Load(); return logging; }
            set { Load(); logging = value; Save("Logging", value ? 1 : 0); }
        }

        /// <summary>
        /// Roll the log once it passes this size. One previous file is kept, so the worst case on
        /// disk is twice this - bounded, and knowable from the setting alone.
        /// </summary>
        public static int LogMaxKB
        {
            get { Load(); return logMaxKB < 64 ? 64 : logMaxKB; }
            set { Load(); logMaxKB = value; Save("LogMaxKB", value); }
        }
    }
}
