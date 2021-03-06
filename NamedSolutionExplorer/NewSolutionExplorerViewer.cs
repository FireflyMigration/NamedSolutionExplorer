﻿//------------------------------------------------------------------------------
// <copyright file="NewSolutionExplorerViewer.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using log4net;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NamedSolutionExplorer.Models;
using Constants = EnvDTE.Constants;
using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
{
    /// <summary>
    ///     Command handler
    /// </summary>
    internal class NewSolutionExplorerViewer
    {
        #region Statics

        /// <summary>
        ///     Command ID.
        /// </summary>
        public const int NewNamedSolutionExplorerCommandId = 0x0100;

        public const int RenameSolutionExplorerCommandId = 0x0100;

        private static readonly ILog _log = LogManager.GetLogger(typeof(NewSolutionExplorerViewer));

        /// <summary>
        ///     Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid NewNamedSolutionExplorerCommandSet =
            new Guid("2910728c-1593-479d-9a07-8ee1fe3e8e55");

        public static readonly Guid RenameSolutionExplorerCommandSet = new Guid("59e02eb3-7904-4165-84f5-51fce173f18d");

        #endregion

        #region Private Vars

        private AsyncPackage _package;

        #endregion

        #region Public Methods

        public async void InitialiseAsync(AsyncPackage package)
        {
            try
            {
                _package = package;
                var commandService =
                    await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
                if (commandService != null)
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    var renameCommandID =
                        new CommandID(RenameSolutionExplorerCommandSet, RenameSolutionExplorerCommandId);
                    var renameMenuItem = new MenuCommand(RenameSolutionExplorerWindow, renameCommandID);
                    commandService.AddCommand(renameMenuItem);

                    var menuCommandID = new CommandID(NewNamedSolutionExplorerCommandSet,
                        NewNamedSolutionExplorerCommandId);
                    var menuItem = new MenuCommand(OpenSolutionExplorerView, menuCommandID);
                    commandService.AddCommand(menuItem);
                }
                else
                {
                    _log.Info("failed to find commandService");
                }
            }
            catch (Exception e)
            {
                _log.Error("InitialiseAsync failed", e);
            }
        }

        #endregion

        #region Private Methods

        private static void activateSolutionExplorerWindow(DTE2 dte)
        {
            dte.Windows.Item(Constants.vsWindowKindSolutionExplorer).Activate();
        }

        private static string getNameFromHierarchyItem(UIHierarchyItem uihItem)
        {
            var name = uihItem.Name;
            var pi = uihItem.Object as ProjectItem;
            if (pi != null && pi.ContainingProject != null)
                name = string.Format("{0} ({1})", name, pi.ContainingProject.Name);

            return name;
        }

        private async Task<T> GetService<T>()
        {
            return (T) await _package.GetServiceAsync(typeof(T));
        }

        private static async Task openNewScopedExplorerWindow(DTE2 dte)
        {
            await Task.Run(async () =>
            {
                var commands = dte.Commands.Cast<Command>();
                activateSolutionExplorerWindow(dte);
                var openSolutionView =
                    commands.FirstOrDefault(
                        x => x.Name == "ProjectandSolutionContextMenus.Project.SolutionExplorer.NewScopedWindow");
                if (openSolutionView != null)
                    if (openSolutionView.IsAvailable)
                    {
                        dte.Commands.Raise(openSolutionView.Guid, openSolutionView.ID, null, null);
                        await waitUntilWindowOpened(dte);
                    }
                    else
                    {
                        _log.Debug("NewScopedWindow not yet available");
                    }
            });
        }

        private void OpenSolutionExplorerView(object sender, EventArgs e)
        {
            Task.Run(async () =>
            {
                try
                {
                    await OpenSolutionExplorerViewAsync();
                }
                catch (Exception ex)
                {
                    _log.Error("Failed to open solution explorer view", ex);
                }
            });
        }

        private async Task OpenSolutionExplorerViewAsync(NamedSolutionExplorerWindowConfig config = null)
        {
            var dte = await GetService<DTE>() as DTE2;

            renamedExistingSolutionExplorerWindows(dte);

            // call the "open visual studio new menu view"
            await openNewScopedExplorerWindow(dte).ContinueWith(async _ =>
            {
                // solutionexplorer automatically points to the last one created
                var windowName = config?.Name;

                await renameCurrentSolutionExplorerWindow(dte,
                    item => string.IsNullOrEmpty(windowName) ? getNameFromHierarchyItem(item) : windowName);
            });
        }

        private async Task<string> PromptUserForNewCaption(Window window)
        {
            var existingCaption = window.Caption;
            var result = await RenameDialogWindow.GetNewCaption(existingCaption);

            return result;
        }

        private async Task renameCurrentSolutionExplorerWindow(DTE2 dte,
            Func<UIHierarchyItem, string> newCaptionFunc)
        {
            var UIH = dte.ToolWindows.SolutionExplorer;
            var wnd = UIH.Parent;
            var UIHItem = UIH.UIHierarchyItems.Item(1);

            var newCaption = newCaptionFunc(UIHItem);
            UIH.Parent.Caption = newCaption;
            await setWindowDockedCaption(wnd, newCaption);
        }

        private static void renamedExistingSolutionExplorerWindows(DTE2 dte)
        {
            var list = new List<Window>();
            foreach (Window w in dte.Windows)
                if (w.Caption == "Solution Explorer")
                    list.Add(w);

            if (list.Count > 1)
                for (var i = 1; i < list.Count; i++)
                    list[i].Caption = $"Solution Explorer {i}";
        }

        private void RenameSolutionExplorerWindow(object sender, EventArgs e)
        {
            Task.Run(async () =>
            {
                try
                {
                    var dte = await GetService<DTE>() as DTE2;

                    var w = dte.ActiveWindow;
                    if (Utilities.IsSolutionExplorer(w))
                    {
                        var newName = await PromptUserForNewCaption(w);

                        if (!string.IsNullOrEmpty(newName)) await setWindowDockedCaption(w, newName);
                    }
                }
                catch (Exception ex)
                {
                    _log.Error("Failed to rename", ex);
                }
            });
        }

        private async Task setWindowDockedCaption(Window wnd, string newCaption)
        {
            var shell = await GetService<IVsUIShell>();

            var frame = Utilities.GetWindowFrameFromWindow(wnd, shell);
            frame.SetProperty((int) __VSFPROPID.VSFPROPID_Caption, newCaption);
        }

        private static async Task waitUntilWindowOpened(DTE2 dte)
        {
            await Task.Delay(new TimeSpan(0, 0, 1));
        }

        #endregion

        #region Nested Types

        private class DontSaveSettings : EventArgs
        {
        }

        #endregion

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