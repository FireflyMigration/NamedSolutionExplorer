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

        #endregion

        #region Properties

        public string HierarchyId { get; set; }
        public string Name { get; set; }

        public SizeAndPosition SizeAndPosition { get; set; }

        #endregion
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

        #endregion

        #region Public Methods

        public async Task ApplyToAsync(Window w, IVsUIShell uiShell)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var frame = Utilities.GetFrame(w, uiShell);
            var relativeTo = Guid.Empty;

            frame.SetFramePos(DockPosition, relativeTo, 0, 0, Left, Top);
        }

        public static async Task<SizeAndPosition> FromWindowAsync(Window w, IVsUIShell uiShell)
        {
            var frame = Utilities.GetFrame(w, uiShell);

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

            ret.DockPosition = (VSSETFRAMEPOS) result;

            return ret;
        }

        #endregion
    }
}