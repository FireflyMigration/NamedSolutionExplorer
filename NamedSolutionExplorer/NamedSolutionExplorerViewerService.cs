using EnvDTE;

using EnvDTE80;

using log4net;

using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;

using System;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
{
    public class NamedSolutionExplorerViewerService
    {
        private NewSolutionExplorerViewer _viewer;
        private ILog _log = LogManager.GetLogger(typeof(NamedSolutionExplorerViewerService));

        public NamedSolutionExplorerViewerService()
        {
            _viewer = new NewSolutionExplorerViewer();
        }

        public async Task InitialiseAsync(AsyncPackage package)
        {
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
                var store = GetWritableSettingsStore();
                if (!store.CollectionExists("NamedSolutionExplorer"))
                {
                    store.CreateCollection("NamedSolutionExplorer");
                }

                if (store.PropertyExists("NamedSolutionExplorer", solutionId))
                {
                    var settingsString = store.GetString("NamedSolutionExplorer", solutionId);

                    var savedSettings = await NSESettings.FromString(settingsString);

                    restoreSettings(savedSettings);
                }
            }
        }

        private string BuildHierarchyPathForProject(ProjectItem projectItem)
        {
            var project = projectItem.ContainingProject;
            Project current = GetParentProject(project);
            string path = project.Name;
            while (current != null)
            {
                path = current.Name + "\\" + path;
                current = GetParentProject(current);
            }

            return path;
        }

        private Project GetParentProject(Project project)
        {
            try
            {
                return project.ParentProjectItem != null
                    ? project.ParentProjectItem.ContainingProject
                    : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void restoreSettings(NSESettings savedSettings)
        {
            foreach (var settin in savedSettings.Settings)
            {
                restoreWindow(settin);
            }
        }

        private void restoreWindow(NSESettings.aNSE settin)
        {
            if (string.IsNullOrEmpty(settin.HierarchyId)) return;

            // select the item in the solutionexplorer window
            var d = GetDTE() as DTE2;
            var se = d.ToolWindows.SolutionExplorer;
            try
            {
                var found = se.GetItem(settin.HierarchyId);
                if (found == null)
                {
                    _log.DebugFormat("Failed to find item {0}", settin.HierarchyId);
                }
                else
                {
                    _log.InfoFormat("Selecting item {0}", settin.HierarchyId);
                    // select item in the explorer
                    found.Select(vsUISelectionType.vsUISelectionTypeSelect);
                    _viewer.OpenSEV();
                    Console.WriteLine(found.Name);
                    // create a new window
                }
            }
            catch (Exception e)
            {
                // ignore, couldnt find the item
            }
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
                store.SetString("NamedSolutionExplorer", solutionId, settings.ToString());
            }
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
            if (!string.IsNullOrEmpty(hierarchyId) && !string.IsNullOrEmpty(name))
                ret.Add(new NSESettings.aNSE(hierarchyId, name));
        }

        private string GetName(Window window)
        {
            return window.Caption;
        }

        private string GetHierarchyId(Window window)
        {
            if (!IsSolutionExplorer(window)) return null;

            UIHierarchy uiHierarchy = (UIHierarchy)window.Object;

            UIHierarchyItem firstItem = uiHierarchy.UIHierarchyItems.Item(1);

            return GetObjectPath(firstItem);

            return window.DocumentData?.ToString();
        }

        private string GetObjectPath(UIHierarchyItem tgt)
        {
            var name = tgt.Name;

            if (tgt.Collection == null)
            {
                return name;
            }

            var parent = tgt.Collection.Parent;
            var sol = parent as Solution;
            var pi = parent as ProjectItem;
            var p = parent as Project;

            if (p != null)
            {
                return GetObjectPath(p) + "\\" + name;
            }

            if (pi != null)
            {
                return GetObjectPath(pi) + "\\" + name;
            }

            var solutionName = GetDTE().Solution.Properties.Item("Name").Value.ToString();

            return string.Format("{0}\\{1}", solutionName, name);
        }

        private string GetObjectPath(ProjectItem pi)
        {
            var ret = pi.Name;

            return GetObjectPath(pi.ContainingProject) + "\\" + ret;
        }

        private string GetObjectPath(Project p)
        {
            return "\\" + p.Name;
        }

        private bool IsSolutionExplorer(Window window)
        {
            return window.ObjectKind == EnvDTE.Constants.vsWindowKindSolutionExplorer;
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