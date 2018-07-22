using EnvDTE;

using log4net;

using Microsoft.VisualStudio.Shell;

using System;
using System.Threading.Tasks;

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

        private readonly NewSolutionExplorerViewer _viewer;

        #endregion Private Vars

        #region Constructors

        public NamedSolutionExplorerViewerService()
        {
            _viewer = new NewSolutionExplorerViewer();
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
            _viewer.InitialiseAsync(package);
        }

        #endregion Public Methods

        #region Private Methods

        private string getNames(string uniqueName)
        {
            return uniqueName.Replace(".csproj", "").Replace(".csProj", "");
        }

        #endregion Private Methods
    }
}