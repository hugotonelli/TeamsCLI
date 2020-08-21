using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;

namespace TeamsCLI.Settings
{
    class SettingsManager
    {
        const string UserSettingsFileName = "usersettings.json";

        private static readonly JsonSerializerOptions Options = new JsonSerializerOptions
        {
            WriteIndented = true,
            IgnoreNullValues = true,
        };

        public static ClientSettings ReadOrCreateSettings()
        {
            string settings;
            if (File.Exists(UserSettingsFileName))
            {
                settings = File.ReadAllText(UserSettingsFileName);
                return JsonSerializer.Deserialize<ClientSettings>(settings, Options);
            }

            var newSettings = new ClientSettings();
            SaveSettings(newSettings);
            return newSettings;
        }

        public static void SaveSettings(ClientSettings newSettings)
        {
            string settings = JsonSerializer.Serialize(newSettings, Options);
            File.WriteAllText(UserSettingsFileName, settings);
        }
    }
}
