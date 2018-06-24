//------------------------------------------------------------------------------
// <copyright file="NewSolutionExplorerViewer.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using EnvDTE;

using EnvDTE80;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Threading.Tasks;

namespace NamedSolutionExplorer
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal class NewSolutionExplorerViewer
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2910728c-1593-479d-9a07-8ee1fe3e8e55");

        private AsyncPackage _package;

        public async void InitialiseAsync(AsyncPackage package)
        {
            _package = package;
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.OpenSolutionExplorerView, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private async Task<T> GetService<T>()
        {
            return (T)await _package.GetServiceAsync(typeof(T));
        }

        private async void setSolutionExplorerToolWindowCaption(string caption)
        {
            IVsWindowFrame frame = null;
            var shell = await GetService<IVsUIShell>();

            shell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fFrameOnly, new Guid(ToolWindowGuids.SolutionExplorer), out frame);

            frame.SetProperty((int)__VSFPROPID.VSFPROPID_Caption, caption);
        }

        public void OpenSEV()
        {
            OpenSolutionExplorerView(null, null);
        }

        private async void OpenSolutionExplorerView(object sender, EventArgs e)
        {
            DTE2 dte = await GetService<DTE>() as DTE2;
            renamedExistingSolutionExplorerWindows(dte);

            // call the "open visual studio new menu view"

            openNewScopedExplorerWindow(dte);

            // solutionexplorer automatically points to the last one created
            renameCurrentSolutionExplorerWindowToFirstItemInList(dte);
            saveSettings();
        }

        private async void saveSettings()
        {
            var svc = await GetService<NamedSolutionExplorerViewerService>();
            await svc.SaveSettings();
        }

        private void renameCurrentSolutionExplorerWindowToFirstItemInList(DTE2 dte)
        {
            UIHierarchy UIH = dte.ToolWindows.SolutionExplorer;

            UIHierarchyItem UIHItem = UIH.UIHierarchyItems.Item(1);
            var newCaption = getWindowName(UIHItem);
            UIH.Parent.Caption = newCaption;
            setSolutionExplorerToolWindowCaption(newCaption);
        }

        private static string getWindowName(UIHierarchyItem uihItem)
        {
            var name = uihItem.Name;
            var pi = uihItem.Object as ProjectItem;
            if (pi != null && pi.ContainingProject != null)
            {
                name = string.Format("{0} ({1})", name, pi.ContainingProject.Name);
            }

            return name;
        }

        private static void openNewScopedExplorerWindow(DTE2 dte)
        {
            var commands = dte.Commands.Cast<Command>();
            var openSolutionView =
                commands.FirstOrDefault(
                    x => x.Name == "ProjectandSolutionContextMenus.Project.SolutionExplorer.NewScopedWindow");
            if (openSolutionView != null)
            {
                dte.Commands.Raise(openSolutionView.Guid, openSolutionView.ID, null, null);
            }
        }

        private static void renamedExistingSolutionExplorerWindows(DTE2 dte)
        {
            List<Window> list = new List<Window>();
            foreach (Window w in dte.Windows)
            {
                if (w.Caption == "Solution Explorer")
                {
                    list.Add(w);
                }
            }

            if (list.Count > 1)
            {
                for (int i = 1; i < list.Count; i++)
                {
                    list[i].Caption = $"Solution Explorer {i}";
                }
            }
        }

        ///// <summary>
        ///// This function is the callback used to execute the command when the menu item is clicked.
        ///// See the constructor to see how the menu item is associated with this function using
        ///// OleMenuCommandService service and MenuCommand class.
        ///// </summary>
        ///// <param name="sender">Event sender.</param>
        ///// <param name="e">Event args.</param>
        //private void MenuItemCallback(object sender, EventArgs e)
        //{
        //    string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
        //    string title = "NewSolutionExplorerViewer";

        //    // Show a message box to prove we were here
        //    VsShellUtilities.ShowMessageBox(
        //        this.ServiceProvider,
        //        message,
        //        title,
        //        OLEMSGICON.OLEMSGICON_INFO,
        //        OLEMSGBUTTON.OLEMSGBUTTON_OK,
        //        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        //}
    }
}