using System;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer.Models
{
    public class NamedSolutionExplorerWindowConfig
    {
        #region Constructors

        public NamedSolutionExplorerWindowConfig(string hierarchyId, string name)
        {
            HierarchyId = hierarchyId;
            Name = name;
        }

        #endregion Constructors

        #region Properties

        public string HierarchyId { get; set; }
        public string Name { get; set; }

        public SizeAndPosition SizeAndPosition { get; set; }

        #endregion Properties
    }

    public class SizeAndPosition
    {
        #region Properties

        public VSSETFRAMEPOS DockPosition { get; set; }
        public int Height { get; set; }
        public bool IsFloating { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public string WindowState { get; set; }

        #endregion Properties

        #region Public Methods

        public async Task ApplyToAsync(Window w, IVsUIShell uiShell)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var frame = GetFrame(w, uiShell);
            var relativeTo = Guid.Empty;

            frame.SetFramePos(DockPosition, relativeTo, 0, 0, Left, Top);
        }

        public static async Task<SizeAndPosition> FromWindowAsync(Window w, IVsUIShell uiShell)
        {
            var frame = GetFrame(w, uiShell);

            Guid pguidRelativeTo;
            int px;
            int py;
            int pcx;
            int pcy
                ;

            var ret = new SizeAndPosition();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var result = frame.GetFramePos(new[]
                {
                    VSSETFRAMEPOS.SFP_fSize
                }, out pguidRelativeTo, out px,
                out py,
                out pcx, out pcy);

            ret.Left = pcx;
            ret.Top = pcy;

            ret.DockPosition = (VSSETFRAMEPOS)result;

            return ret;
        }

        #endregion Public Methods

        #region Private Methods

        private static IVsWindowFrame GetFrame(Window w, IVsUIShell uiShell)
        {
            return GetWindowFrameFromWindow(w, uiShell);
        }

        private static IVsWindowFrame GetWindowFrameFromGuid(Guid guid, IVsUIShell uiShell)
        {
            var slotGuid = guid;
            IVsWindowFrame wndFrame;

            uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fFrameOnly, ref slotGuid, out wndFrame);
            return wndFrame;
        }

        private static IVsWindowFrame GetWindowFrameFromWindow(Window window, IVsUIShell uiShell)
        {
            if (window == null)
                return null;
            if (window.ObjectKind == null || window.ObjectKind == string.Empty)
                return null;
            return GetWindowFrameFromGuid(new Guid(window.ObjectKind), uiShell);
        }

        #endregion Private Methods
    }
}