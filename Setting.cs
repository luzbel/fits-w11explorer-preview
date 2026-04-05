using System;
using Microsoft.Win32;
using System.Diagnostics;

namespace FitsPreviewHandler
{
    /// <summary>
    /// Safe registry configuration for the FITS Preview Handler.
    /// Operates correctly within the restricted Low-Integrity (Low-IL) prevhost.exe environment.
    /// </summary>
    public static class Settings
    {
        public const string REG_PATH = @"Software\AppDataLow\FitsPreviewHandler";
        public const string VAL_SHOW_IMAGE = "ShowImage";
        public const string VAL_ENABLE_LOG = "EnableTracing";
        public const string VAL_SPLITTER_POS = "SplitterDistance";

        public static bool ShowImage => ReadBool(VAL_SHOW_IMAGE, true);
        public static bool EnableTracing => ReadBool(VAL_ENABLE_LOG, false);

        public static int SplitterDistance
        {
            get 
            {
                int val = ReadInt(VAL_SPLITTER_POS, -1);
                if (val > 0 && val < 250) return 250; 
                return val;
            }
            set => WriteInt(VAL_SPLITTER_POS, value);
        }

        public static void Refresh() { }

        private static bool ReadBool(string valueName, bool defaultValue)
        {
            try
            {
                // 1. Try HKCU (User override)
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_PATH, false))
                {
                    if (key != null)
                    {
                        object valObj = key.GetValue(valueName);
                        if (valObj is int intVal) return intVal != 0;
                    }
                }

                // 2. Try HKLM (Global default)
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(REG_PATH, false))
                {
                    if (key != null)
                    {
                        object valObj = key.GetValue(valueName);
                        if (valObj is int intVal) return intVal != 0;
                    }
                }
            }
            catch (Exception ex)
            {
                // Fall-safe: if registry is unreachable (security), return default
                Debug.WriteLine($"[Settings] Failed to read {valueName} from registry: {ex.Message}");
            }
            return defaultValue;
        }
        private static int ReadInt(string valueName, int defaultValue)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_PATH, false))
                {
                    if (key != null)
                    {
                        object valObj = key.GetValue(valueName);
                        if (valObj is int intVal) return intVal;
                    }
                }
            } catch (Exception ex) { Debug.WriteLine($"[Settings] ReadInt HKCU Error: {ex.Message}"); }

            try {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(REG_PATH, false))
                {
                    if (key != null)
                    {
                        object valObj = key.GetValue(valueName);
                        if (valObj is int intVal) return intVal;
                    }
                }
            } catch (Exception ex) { Debug.WriteLine($"[Settings] ReadInt HKLM Error: {ex.Message}"); }

            return defaultValue;
        }

        private static void WriteInt(string valueName, int value)
        {
            try
            {
                // Note: Writing to AppDataLow in registry normally works in Low-IL (prevhost)
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_PATH))
                {
                    if (key != null) key.SetValue(valueName, value, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Settings] Failed to write {valueName}: {ex.Message}");
            }
        }
    }
}
