using EnvDTE;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;

using System;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;

namespace NamedSolutionExplorer
{
    public class NamedSolutionExplorerWindowConfig
    {
        public NamedSolutionExplorerWindowConfig(string hierarchyId, string name)
        {
            HierarchyId = hierarchyId;
            Name = name;
        }

        public string HierarchyId { get; set; }
        public string Name { get; set; }

        public SizeAndPosition SizeAndPosition { get; set; }
    }

    public class SizeAndPosition
    {
        public bool IsFloating { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string WindowState { get; set; }

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

        private static IVsWindowFrame GetWindowFrameFromGuid(Guid guid, IVsUIShell uiShell)
        {
            Guid slotGuid = guid;
            IVsWindowFrame wndFrame;

            uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fFrameOnly, ref slotGuid, out wndFrame);
            return wndFrame;
        }

        private static IVsWindowFrame GetWindowFrameFromWindow(EnvDTE.Window window, IVsUIShell uiShell)
        {
            if (window == null)
                return null;
            if (window.ObjectKind == null || window.ObjectKind == String.Empty)
                return null;
            return GetWindowFrameFromGuid(new Guid(window.ObjectKind), uiShell);
        }

        private static IVsWindowFrame GetFrame(Window w, IVsUIShell uiShell)
        {
            return GetWindowFrameFromWindow(w, uiShell);
        }

        public async Task ApplyToAsync(Window w, IVsUIShell uiShell)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var frame = GetFrame(w, uiShell);
            var relativeTo = Guid.Empty;

            frame.SetFramePos(this.DockPosition, relativeTo, 0, 0, this.Left, this.Top);
        }

        public VSSETFRAMEPOS DockPosition { get; set; }
    }
}