using EnvDTE;

using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;

using Newtonsoft.Json;

using System.Collections.Generic;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;

namespace ContextMenuOnSolutionExplorer
{
    public class NSESettings
    {
        public class aNSE
        {
            public string HierarchyId { get; set; }
            public string Name { get; set; }
        }

        public static async Task<NSESettings> FromString(string src)
        {
            return await Task.Run(() => JsonConvert.DeserializeObject<NSESettings>(src));
        }

        public List<aNSE> Settings { get; set; } = new List<aNSE>();

        public void Add(aNSE src)
        {
            Settings.Add(src);
        }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public class NamedSolutionExplorerViewerService
    {
        public NamedSolutionExplorerViewerService(Package package)
        {
            NewSolutionExplorerViewer.Initialize(package);
        }

        private WritableSettingsStore GetWritableSettingsStore()
        {
            var shellSettingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
            return shellSettingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
        }

        public async Task LoadAndApplySettings()
        {
            var solutionId = await GetSolutionIdentifier();
            var store = GetWritableSettingsStore();
            var settingsString = store.GetString("NamedSolutionExplorer", solutionId);

            var savedSettings = await NSESettings.FromString(settingsString);
        }

        public async Task SaveSettings()
        {
            var solutionId = await GetSolutionIdentifier();
            var store = GetWritableSettingsStore();

            var settings = await CreateSettings();
            store.SetString("NamedSolutionExplorer", solutionId, settings.ToString());
        }

        private async Task<NSESettings> CreateSettings()
        {
            return await Task.Run<NSESettings>(() =>
            {
                var ret = new NSESettings();

                // get all NSE windows
                var dte = GetDTE();
                foreach (Window w in dte.Windows)
                {
                }
                return ret;
            });
        }

        private async Task<string> GetSolutionIdentifier()
        {
            return await Task.Run<string>(() =>
            {
                var dte = GetDTE();

                return dte.Solution.FullName;
            });
        }

        private DTE GetDTE()
        {
            var dte = (DTE)Package.GetGlobalService(typeof(DTE));
            if (dte == null) return null;
            return dte;
        }
    }
}