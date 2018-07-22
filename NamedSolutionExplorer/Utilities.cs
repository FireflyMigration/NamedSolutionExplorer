using System;
using EnvDTE;
using Microsoft.VisualStudio.Shell.Interop;
using Constants = EnvDTE.Constants;

namespace NamedSolutionExplorer
{
    public static class Utilities
    {
        #region Public Methods

        public static IVsWindowFrame GetFrame(Window w, IVsUIShell uiShell)
        {
            return GetWindowFrameFromWindow(w, uiShell);
        }

        public static IVsWindowFrame GetWindowFrameFromGuid(Guid guid, IVsUIShell uiShell)
        {
            var slotGuid = guid;
            IVsWindowFrame wndFrame;

            uiShell.FindToolWindow((uint) __VSFINDTOOLWIN.FTW_fFrameOnly, ref slotGuid, out wndFrame);
            return wndFrame;
        }

        public static IVsWindowFrame GetWindowFrameFromWindow(Window window, IVsUIShell uiShell)
        {
            if (window == null)
                return null;
            if (window.ObjectKind == null || window.ObjectKind == string.Empty)
                return null;
            return GetWindowFrameFromGuid(new Guid(window.ObjectKind), uiShell);
        }

        public static bool IsSolutionExplorer(Window window)
        {
            return window.ObjectKind == Constants.vsWindowKindSolutionExplorer;
        }

        #endregion
    }
}