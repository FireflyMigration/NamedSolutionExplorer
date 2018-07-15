using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using log4net;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NamedSolutionExplorer.Models;
using NamedSolutionExplorer.Repositories;
using static Microsoft.VisualStudio.Shell.ServiceProvider;
using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
{
    /// <summary>
    ///     Responsible for responding to IDE events and coordinating activities for the service
    /// </summary>
    public class NamedSolutionExplorerViewerService
    {
        #region Private Vars

        private readonly ILog _log = LogManager.GetLogger(typeof(NamedSolutionExplorerViewerService));
        private readonly SettingsRepository _repository;
        private AsyncPackage _serviceProvider;
        private IVsSolution _solutionService;
        private readonly NewSolutionExplorerViewer _viewer;

        #endregion Private Vars

        #region Constructors

        public NamedSolutionExplorerViewerService()
        {
            _viewer = new NewSolutionExplorerViewer();
            _repository = new SettingsRepository();
        }

        #endregion Constructors

        #region Public Methods

        public async Task<UIHierarchyItem> FindItem(UIHierarchy solutionExplorerWindow, string uniqueName)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            try
            {
                var names = getNames(uniqueName);

                var item = solutionExplorerWindow.GetItem(names);

                return item;
            }
            catch (ArgumentException)
            {
                _log.Info("Couldn't restore item " + uniqueName + " as couldnt find it in explorer");
            }
            catch (Exception e)
            {
                _log.Error("Error finding item " + uniqueName, e);
            }

            return null;
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

        public async Task SaveSettings()
        {
            await AddSettingsToRepository();
            var settingsFileName = await GetSettingsFilePath();
            await _repository.Save(settingsFileName);
        }

        #endregion Public Methods

        #region Private Methods

        private async Task AddSettingsToRepository()
        {
            // get all NSE windows
            var dte = await GetDTE();
            var uiShell = (IVsUIShell)await GetGlobalServiceAsync(typeof(IVsUIShell));

            foreach (Window w in dte.Windows)
                if (Utilities.IsSolutionExplorer(w))
                    await AddSolutionExplorerAsync(w, uiShell);
        }

        private async Task AddSolutionExplorerAsync(Window window, IVsUIShell uiShell)
        {
            var config = await CreateConfigAsync(window, uiShell);
            if (config != null) _repository.AddOrReplace(config);
        }

        private async Task<NamedSolutionExplorerWindowConfig> CreateConfigAsync(Window window, IVsUIShell uiShell)
        {
            var hierarchyId = GetRootObjectIdentifier(window);

            var name = GetName(window);
            if (!string.IsNullOrEmpty(hierarchyId) && !string.IsNullOrEmpty(name))
            {
                var ret = new NamedSolutionExplorerWindowConfig(hierarchyId, name);
                ret.SizeAndPosition = await GetSizeAndPositionAsync(window, uiShell);

                return ret;
            }

            return null;
        }

        private UIHierarchy GetAndActivateOriginalSolutionExplorer(DTE2 dte2)
        {
            var ret = (UIHierarchy)dte2.ToolWindows.GetToolWindow("Solution Explorer");

            ret.Parent.Activate();

            return ret;
        }

        private async Task<DTE> GetDTE()
        {
            return (DTE)await _serviceProvider.GetServiceAsync(typeof(DTE));
        }

        private string GetRootObjectIdentifier(Window window)
        {
            if (!Utilities.IsSolutionExplorer(window)) return null;

            var uiHierarchy = (UIHierarchy)window.Object;

            var firstItem = uiHierarchy.UIHierarchyItems.Item(1);
            var d2 = (DTE2)GetDTE().Result;
            return GetObjectIdentifier(d2, firstItem);
        }

        private string GetName(Window window)
        {
            return window.Caption;
        }

        private string getNames(string uniqueName)
        {
            return uniqueName.Replace(".csproj", "").Replace(".csProj", "");
        }

        private string GetObjectIdentifier(DTE2 dte2, UIHierarchyItem selectedUIHierarchyItem)
        {
            var name = selectedUIHierarchyItem.Name;

            if (selectedUIHierarchyItem.Collection == null) return name;

            if (selectedUIHierarchyItem.Object is Project)
            {
                Debug.WriteLine("Project node is selected: " + selectedUIHierarchyItem.Name);
                var project = selectedUIHierarchyItem.Object as Project;
                var p = getPath(dte2, project);
                return p;
            }

            if (selectedUIHierarchyItem.Object is ProjectItem)
            {
                Debug.WriteLine("Project item node is selected: " + selectedUIHierarchyItem.Name);
                var pi = selectedUIHierarchyItem.Object as ProjectItem;
                ;
            }
            else if (selectedUIHierarchyItem.Object is Solution)
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

        private async Task<string> GetSettingsFilePath()
        {
            var dte = await GetDTE();
            var settingsFileName = Path.Combine(Path.GetDirectoryName(dte.Solution.FileName),
                "namedsolutionexplorers.json");
            return settingsFileName;
        }

        private async Task<SizeAndPosition> GetSizeAndPositionAsync(Window window, IVsUIShell uiShell)
        {
            var ret = await SizeAndPosition.FromWindowAsync(window, uiShell);
            return ret;
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

                    await _viewer.OpenSolutionExplorerViewAsync(windowConfig);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ARGH!" + e);
                // ignore, couldnt find the item
            }
        }

        private async Task restoreWindows()
        {
            foreach (var config in _repository.WindowConfigs) await restoreWindow(config);
        }

        #endregion Private Methods
    }
}