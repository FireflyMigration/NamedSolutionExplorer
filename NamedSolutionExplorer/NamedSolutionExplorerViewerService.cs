using EnvDTE;

using EnvDTE80;

using log4net;

using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell.Settings;

using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;
using Window = EnvDTE.Window;

namespace NamedSolutionExplorer
{
    public class NamedSolutionExplorerViewerService
    {
        private NewSolutionExplorerViewer _viewer;
        private ILog _log = LogManager.GetLogger(typeof(NamedSolutionExplorerViewerService));
        private IVsSolution _solutionService;

        public NamedSolutionExplorerViewerService()
        {
            _viewer = new NewSolutionExplorerViewer();
        }

        public async Task InitialiseAsync(AsyncPackage package)
        {
            _solutionService = (IVsSolution)await ServiceProvider.GetGlobalServiceAsync(typeof(IVsSolution));
            _viewer.InitialiseAsync(package);
        }

        private WritableSettingsStore GetWritableSettingsStore()
        {
            var shellSettingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
            return shellSettingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
        }

        public async Task LoadAndApplySettings()
        {
            var solutionId = await GetSolutionIdentifier();
            if (!string.IsNullOrEmpty(solutionId))
            {
                var filepath = GetSettingsFilePath();
                if (File.Exists(filepath))
                {
                    var settingsString = File.ReadAllText(filepath);
                    var savedSettings = await NSESettings.FromString(settingsString);

                    await restoreSettings(savedSettings);
                }
            }
        }

        private async Task restoreSettings(NSESettings savedSettings)
        {
            foreach (var settin in savedSettings.Settings)
            {
                await restoreWindow(settin);
            }
        }

        private async Task restoreWindow(NSESettings.aNSE settin)
        {
            if (string.IsNullOrEmpty(settin.HierarchyId)) return;

            // select the item in the solutionexplorer window
            var d = GetDTE() as DTE2;

            var se = getOriginalSolutionExplorer(d);
            var w = se.Parent;
            w.Activate();

            try
            {
                var found = await FindItem(se, settin.HierarchyId);

                if (found == null)
                {
                    _log.DebugFormat("Failed to find item {0}", settin.HierarchyId);
                }
                else
                {
                    _log.InfoFormat("Selecting item {0}", settin.HierarchyId);
                    found.Select(vsUISelectionType.vsUISelectionTypeSelect);

                    await _viewer.OpenSolutionExplorerViewAsync();
                    _log.Debug("Found item " + found.Name);
                    // create a new window
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ARGH!" + e.ToString());
                // ignore, couldnt find the item
            }
        }

        private UIHierarchy getOriginalSolutionExplorer(DTE2 dte2)
        {
            return (UIHierarchy)dte2.ToolWindows.GetToolWindow("Solution Explorer");
        }

        public async Task SaveSettings()
        {
            var solutionId = await GetSolutionIdentifier();
            if (!string.IsNullOrEmpty(solutionId))
            {
                var store = GetWritableSettingsStore();
                if (!store.CollectionExists("NamedSolutionExplorer"))
                {
                    store.CreateCollection("NamedSolutionExplorer");
                }

                var settings = await CreateSettings();
                string settingsFileName = GetSettingsFilePath();
                File.WriteAllText(settingsFileName, settings.ToString());
            }
        }

        private string GetSettingsFilePath()
        {
            var dte = GetDTE();
            var settingsFileName = Path.Combine(Path.GetDirectoryName(dte.Solution.FileName),
                "namedsolutionexplorers.json");
            return settingsFileName;
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
                    if (Utilities.IsSolutionExplorer(w))
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
            if (!string.IsNullOrEmpty(hierarchyId) && !string.IsNullOrEmpty(name))
                ret.Add(new NSESettings.aNSE(hierarchyId, name));
        }

        private string GetName(Window window)
        {
            return window.Caption;
        }

        private string GetHierarchyId(Window window)
        {
            if (!Utilities.IsSolutionExplorer(window)) return null;

            UIHierarchy uiHierarchy = (UIHierarchy)window.Object;

            UIHierarchyItem firstItem = uiHierarchy.UIHierarchyItems.Item(1);
            var d2 = GetDTE() as DTE2;
            return GetObjectIdentifier(d2, firstItem);
        }

        public async Task<UIHierarchyItem> FindItem(UIHierarchy solutionExplorerWindow, string uniqueName)
        {
            IVsHierarchy projectItem = null;
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            try
            {
                var d2 = GetDTE() as DTE2;
                var names = getNames(uniqueName);

                var item = solutionExplorerWindow.GetItem(names);

                return item;
            }
            catch (Exception e)
            {
                _log.Error("Error finding item " + uniqueName, e);
            }

            return null;
        }

        private string getNames(string uniqueName)
        {
            return uniqueName.Replace(".csproj", "").Replace(".csProj", "");
        }

        private string GetObjectIdentifier(DTE2 dte2, UIHierarchyItem tgt)
        {
            var name = tgt.Name;

            if (tgt.Collection == null)
            {
                return name;
            }

            var selectedUIHierarchyItem = tgt;

            if (tgt.Object is EnvDTE.Project)
            {
                Debug.WriteLine("Project node is selected: " + selectedUIHierarchyItem.Name);
                var project = tgt.Object as Project;
                var p = getPath(dte2, project);
                return p;
            }
            else if (selectedUIHierarchyItem.Object is EnvDTE.ProjectItem)
            {
                Debug.WriteLine("Project item node is selected: " + selectedUIHierarchyItem.Name);
                var pi = selectedUIHierarchyItem.Object as ProjectItem;
                ;
            }
            else if (selectedUIHierarchyItem.Object is EnvDTE.Solution)
            {
                Debug.WriteLine("Solution node is selected: " + selectedUIHierarchyItem.Name);
                return null;
            }
            Debug.WriteLine("Couldn't identify node type");
            return null;
        }

        private string getPath(DTE2 dte2, Project project)
        {
            var solutionName = dte2.Solution.Properties.Item("Name").Value.ToString();
            var projectName = project.Name;

            return $"{solutionName}\\{projectName}";
        }

        private string GetObjectIdentifier(ProjectItem pi)
        {
            var ret = pi.Name;

            return GetObjectIdentifier(pi.ContainingProject) + "\\" + ret;
        }

        private string GetObjectIdentifier(Project p)
        {
            return "\\" + p.Name;
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

    public static class Utilities
    {
        public static bool IsSolutionExplorer(Window window)
        {
            return window.ObjectKind == EnvDTE.Constants.vsWindowKindSolutionExplorer;
        }
    }
}