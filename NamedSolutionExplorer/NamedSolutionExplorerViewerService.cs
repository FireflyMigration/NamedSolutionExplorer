using EnvDTE;

using EnvDTE80;

using log4net;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

using static Microsoft.VisualStudio.Shell.ServiceProvider;

using Task = System.Threading.Tasks.Task;
using Window = EnvDTE.Window;

namespace NamedSolutionExplorer
{
    public class NamedSolutionExplorerViewerService
    {
        private NewSolutionExplorerViewer _viewer;
        private ILog _log = LogManager.GetLogger(typeof(NamedSolutionExplorerViewerService));
        private IVsSolution _solutionService;
        private SettingsRepository _repository;
        private AsyncPackage _serviceProvider;

        public NamedSolutionExplorerViewerService()
        {
            _viewer = new NewSolutionExplorerViewer();
            _repository = new SettingsRepository();
        }

        public async Task InitialiseAsync(AsyncPackage package)
        {
            _solutionService = (IVsSolution)await GetGlobalServiceAsync(typeof(IVsSolution));
            _serviceProvider = package;

            _viewer.InitialiseAsync(package);
        }

        public async Task LoadAndApplySettings()
        {
            var filepath = await GetSettingsFilePath();
            if (File.Exists(filepath))
            {
                await _repository.Load(filepath);
                await restoreWindows();
            }
        }

        private async Task restoreWindows()
        {
            foreach (var config in _repository.WindowConfigs)
            {
                await restoreWindow(config);
            }
        }

        private async Task restoreWindow(NamedSolutionExplorerWindowConfig windowConfig)
        {
            if (string.IsNullOrEmpty(windowConfig.HierarchyId)) return;

            // select the item in the solutionexplorer window
            var d = (DTE2)await GetDTE();

            var solutionExplorer = GetAndActivateOriginalSolutionExplorer(d);

            try
            {
                var found = await FindItem(solutionExplorer, windowConfig.HierarchyId);

                if (found == null)
                {
                    _log.DebugFormat("Failed to find item {0}", windowConfig.HierarchyId);
                }
                else
                {
                    _log.InfoFormat("Selecting item {0}", windowConfig.HierarchyId);
                    found.Select(vsUISelectionType.vsUISelectionTypeSelect);

                    await _viewer.OpenSolutionExplorerViewAsync(windowConfig.Name);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ARGH!" + e.ToString());
                // ignore, couldnt find the item
            }
        }

        private UIHierarchy GetAndActivateOriginalSolutionExplorer(DTE2 dte2)
        {
            var ret = (UIHierarchy)dte2.ToolWindows.GetToolWindow("Solution Explorer");

            ret.Parent.Activate();

            return ret;
        }

        public async Task SaveSettings()
        {
            await AddSettingsToRepository();
            var settingsFileName = await GetSettingsFilePath();
            await _repository.Save(settingsFileName);
        }

        private async Task<string> GetSettingsFilePath()
        {
            var dte = await GetDTE();
            var settingsFileName = Path.Combine(Path.GetDirectoryName(dte.Solution.FileName),
                "namedsolutionexplorers.json");
            return settingsFileName;
        }

        private async Task AddSettingsToRepository()
        {
            // get all NSE windows
            var dte = await GetDTE();

            foreach (Window w in dte.Windows)
            {
                if (Utilities.IsSolutionExplorer(w))
                {
                    AddSolutionExplorer(w);
                }
            }
        }

        private void AddSolutionExplorer(Window window)
        {
            var hierarchyId = GetHierarchyId(window);
            var name = GetName(window);
            if (!string.IsNullOrEmpty(hierarchyId) && !string.IsNullOrEmpty(name))
                _repository.AddOrReplace(new NamedSolutionExplorerWindowConfig(hierarchyId, name));
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
            var d2 = (DTE2)GetDTE().Result;
            return GetObjectIdentifier(d2, firstItem);
        }

        public async Task<UIHierarchyItem> FindItem(UIHierarchy solutionExplorerWindow, string uniqueName)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            try
            {
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

        private string GetObjectIdentifier(DTE2 dte2, UIHierarchyItem selectedUIHierarchyItem)
        {
            var name = selectedUIHierarchyItem.Name;

            if (selectedUIHierarchyItem.Collection == null)
            {
                return name;
            }

            if (selectedUIHierarchyItem.Object is EnvDTE.Project)
            {
                Debug.WriteLine("Project node is selected: " + selectedUIHierarchyItem.Name);
                var project = selectedUIHierarchyItem.Object as Project;
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

        private async Task<DTE> GetDTE()
        {
            return (DTE)await _serviceProvider.GetServiceAsync(typeof(DTE));
        }
    }
}