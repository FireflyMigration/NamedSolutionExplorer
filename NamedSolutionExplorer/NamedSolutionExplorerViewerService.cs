using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
{
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
                    if (IsSolutionExplorer(w))
                    {
                        AddSolutionExplorer(ret, w);
                    }
                }
                return ret;
            });
        }

        private void AddSolutionExplorer(NSESettings ret, Window window)
        {
            var hierarchyId = GetHierarchyId(window);
            var name = GetName(window);

            ret.Add(new NSESettings.aNSE(hierarchyId, name));
        }

        private string GetName(Window window)
        {
            return window.Caption;
        }

        private string GetHierarchyId(Window window)
        {
            return window.DocumentData.ToString();
        }

        private bool IsSolutionExplorer(Window window)
        {
            return true;
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